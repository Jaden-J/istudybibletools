using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BibleCommon.Services;
using BibleCommon.Common;

namespace BibleConfigurator
{
    public class SupplementalBibleForm: BaseSupplementalForm
    {
        public SupplementalBibleForm(Microsoft.Office.Interop.OneNote.Application oneNoteApp, MainForm form)
            : base(oneNoteApp, form)
        { }        

        protected override string GetValidSupplementalNotebookId()
        {
            return SettingsManager.Instance.GetValidSupplementalBibleNotebookId(OneNoteApp, true);
        }        

        protected override int GetSupplementalModulesCount()
        {
            return SettingsManager.Instance.SupplementalBibleModules.Count;
        }

        protected override string GetSupplementalModuleName(int index)
        {
            return SettingsManager.Instance.SupplementalBibleModules[index];
        }

        protected override string FormDescription
        {
            get
            {
                return 
@"Данная форма предназначена для управления справочной Библией. 

Обратите внимание:
  - в справочную Библию можно добавлять только модули версии 2.0 и выше;
  - добавление нового модуля в справочную Библию может занимать несколько часов.  
";
            }
        }

        protected override List<string> CommitChanges(ModuleInfo selectedModuleInfo)
        {
            BibleParallelTranslationConnectionResult result;

            Dictionary<string, string> strongTermLinksCache = null;

            if (SettingsManager.Instance.SupplementalBibleModules.Count > 0)
            {
                int stagesCount = selectedModuleInfo.Type == ModuleType.Strong ? 2 : 1;

                if (selectedModuleInfo.Type == ModuleType.Strong)
                {
                    DictionaryManager.AddDictionary(OneNoteApp, selectedModuleInfo.ShortName, FolderBrowserDialog.SelectedPath, true);
                    strongTermLinksCache = RunIndexStrong(selectedModuleInfo, 1, stagesCount);
                }

                int chaptersCount = ModulesManager.GetBibleChaptersCount(SettingsManager.Instance.SupplementalBibleModules.First(), false);
                MainForm.PrepareForExternalProcessing(chaptersCount, 1, BibleCommon.Resources.Constants.AddParallelBibleTranslationStart);
                string stagesString = stagesCount == 1 ? string.Empty : string.Format("{0} {1}/{1}: ", BibleCommon.Resources.Constants.Stage, stagesCount);
                Logger.Preffix = string.Format("{0}{1}: ", stagesString, BibleCommon.Resources.Constants.AddParallelBibleTranslation); 
                BibleCommon.Services.Logger.LogMessage(Logger.Preffix);
                result = SupplementalBibleManager.AddParallelBible(OneNoteApp, selectedModuleInfo.ShortName, strongTermLinksCache, Logger);

                MainForm.ExternalProcessingDone(BibleCommon.Resources.Constants.AddParallelBibleTranslationFinishMessage);
            }
            else
            {
                int stagesCount = selectedModuleInfo.Type == ModuleType.Strong ? 3 : 2;

                int chaptersCount = ModulesManager.GetBibleChaptersCount(selectedModuleInfo.ShortName, false);
                MainForm.PrepareForExternalProcessing(chaptersCount, 1, BibleCommon.Resources.Constants.CreateSupplementalBibleStart);
                Logger.Preffix = string.Format("{0} 1/{1}: {2}: ", BibleCommon.Resources.Constants.Stage, stagesCount, BibleCommon.Resources.Constants.CreateSupplementalBible);
                BibleCommon.Services.Logger.LogMessage(Logger.Preffix);

                if (selectedModuleInfo.Type == ModuleType.Strong)
                    DictionaryManager.AddDictionary(OneNoteApp, selectedModuleInfo.ShortName, FolderBrowserDialog.SelectedPath, false);                
                SupplementalBibleManager.CreateSupplementalBible(OneNoteApp, selectedModuleInfo.ShortName, FolderBrowserDialog.SelectedPath, Logger);
                
                if (selectedModuleInfo.Type == ModuleType.Strong)                
                    strongTermLinksCache = RunIndexStrong(selectedModuleInfo, 2, stagesCount);                                    

                MainForm.PrepareForExternalProcessing(chaptersCount, 1, BibleCommon.Resources.Constants.LinkSupplementalBibleStart);
                Logger.Preffix = string.Format("{0} {1}/{1}: {2}: ", BibleCommon.Resources.Constants.Stage, stagesCount, BibleCommon.Resources.Constants.LinkSupplementalBible);
                BibleCommon.Services.Logger.LogMessage(Logger.Preffix);
                result = SupplementalBibleManager.LinkSupplementalBibleWithMainBible(OneNoteApp, 0, strongTermLinksCache, Logger);

                MainForm.ExternalProcessingDone(BibleCommon.Resources.Constants.CreateSupplementalBibleFinish);
            }

            return result.Errors.ConvertAll(ex => ex.Message);
        }

        private Dictionary<string, string> RunIndexStrong(ModuleInfo moduleInfo, int stage, int stagesCount)
        {
            int strongTermsCount = moduleInfo.DictionaryTermsCount.GetValueOrDefault(14700);
            MainForm.PrepareForExternalProcessing(strongTermsCount, 1, BibleCommon.Resources.Constants.IndexStrongDictionaryStart);
            Logger.Preffix = string.Format("{0} {1}/{2}: {3}: ", BibleCommon.Resources.Constants.Stage, stage, stagesCount, BibleCommon.Resources.Constants.IndexStrongDictionary);
            BibleCommon.Services.Logger.LogMessage(Logger.Preffix);

            if (DictionaryTermsCacheManager.CacheIsActive(moduleInfo.ShortName))
                return DictionaryTermsCacheManager.LoadCachedDictionary(moduleInfo.ShortName);
            else
                return DictionaryTermsCacheManager.GenerateCache(OneNoteApp, moduleInfo, Logger);
        }

        protected override bool CanModuleBeDeleted(ModuleInfo moduleInfo, int index)
        {
            return index != 0 || index == GetSupplementalModulesCount() - 1;
        }

        protected override void DeleteModule(string moduleShortName)
        {
            int chaptersCount = ModulesManager.GetBibleChaptersCount(moduleShortName, false);
            MainForm.PrepareForExternalProcessing(chaptersCount, 1, BibleCommon.Resources.Constants.RemoveParallelBibleTranslation);
            var removeResult = SupplementalBibleManager.RemoveSupplementalBibleModule(OneNoteApp, moduleShortName, Logger);
            MainForm.ExternalProcessingDone(
                removeResult == SupplementalBibleManager.RemoveResult.RemoveModule
                    ? BibleCommon.Resources.Constants.RemoveParallelBibleTranslationFinishMessage
                    : BibleCommon.Resources.Constants.RemoveSupplementalBibleFinishMessage);
        }

        protected override string CloseSupplementalNotebookQuestionText
        {
            get { return BibleCommon.Resources.Constants.DeleteSupplementalBibleQuestion; }
        }

        protected override void CloseSupplementalNotebook()
        {
            SupplementalBibleManager.CloseSupplementalBible(OneNoteApp);
        }

        protected override bool IsModuleSupported(ModuleInfo moduleInfo)
        {
            return BibleParallelTranslationManager.IsModuleSupported(moduleInfo) 
                && (moduleInfo.Type == ModuleType.Bible || moduleInfo.Type == ModuleType.Strong);
        }

        protected override bool SupplementalModuleAlreadyAdded(string moduleShortName)
        {
            return SettingsManager.Instance.SupplementalBibleModules.Contains(moduleShortName);
        }

        protected override string GetFormText()
        {
            return BibleCommon.Resources.Constants.SupplementalBibleManagement;
        }

        protected override string GetChkUseText()
        {
            return BibleCommon.Resources.Constants.UseSupplementalBible;
        }

        protected override bool IsBaseModuleSupported()
        {
            return BibleParallelTranslationManager.IsModuleSupported(SettingsManager.Instance.CurrentModule);
        }

        protected override string DeleteModuleQuestionText
        {
            get { return BibleCommon.Resources.Constants.DeleteThisModuleFromSupplementalBibleQuestion; }
        }

        protected override bool CanModuleBeAdded(ModuleInfo moduleInfo)
        {
            return true;
        }

        protected override bool CanNotebookBeClosed()
        {
            return true;
        }

        protected override string NotebookCannotBeClosedText
        {
            get { return string.Empty; }
        }
    }
}
