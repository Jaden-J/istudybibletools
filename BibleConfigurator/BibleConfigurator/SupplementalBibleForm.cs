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

        protected override void ClearSupplementalModulesInSettingsStorage()
        {
            SettingsManager.Instance.SupplementalBibleModules.Clear(); 
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

            if (SettingsManager.Instance.SupplementalBibleModules.Count > 0)
            {
                int chaptersCount = ModulesManager.GetBibleChaptersCount(SettingsManager.Instance.SupplementalBibleModules.First());
                MainForm.PrepareForExternalProcessing(chaptersCount, 1, BibleCommon.Resources.Constants.AddParallelBibleTranslation);
                result = SupplementalBibleManager.AddParallelBible(OneNoteApp, selectedModuleInfo.ShortName, FolderBrowserDialog.SelectedPath, Logger);
                MainForm.ExternalProcessingDone(BibleCommon.Resources.Constants.AddParallelBibleTranslationFinishMessage);
            }
            else
            {
                int stagesCount = selectedModuleInfo.Type == ModuleType.Strong ? 3 : 2;

                int chaptersCount = ModulesManager.GetBibleChaptersCount(selectedModuleInfo.ShortName);
                MainForm.PrepareForExternalProcessing(chaptersCount, 1, BibleCommon.Resources.Constants.CreateSupplementalBibleStart);
                Logger.Preffix = string.Format("{0} 1/{1}: {2}: ", BibleCommon.Resources.Constants.Stage, stagesCount, BibleCommon.Resources.Constants.CreateSupplementalBible);
                BibleCommon.Services.Logger.LogMessage(Logger.Preffix);
                SupplementalBibleManager.CreateSupplementalBible(OneNoteApp, selectedModuleInfo.ShortName, FolderBrowserDialog.SelectedPath, Logger);

                Dictionary<string, string> strongTermLinksCache = null;
                if (selectedModuleInfo.Type == ModuleType.Strong)
                {
                    int strongTermsCount = !string.IsNullOrEmpty(selectedModuleInfo.StrongNumbersCount) ? int.Parse(selectedModuleInfo.StrongNumbersCount) : 14700;
                    MainForm.PrepareForExternalProcessing(strongTermsCount, 1, BibleCommon.Resources.Constants.IndexStrongDictionaryStart);
                    Logger.Preffix = string.Format("{0} 2/{1}: {2}: ", BibleCommon.Resources.Constants.Stage, stagesCount, BibleCommon.Resources.Constants.IndexStrongDictionary);
                    BibleCommon.Services.Logger.LogMessage(Logger.Preffix);
                    strongTermLinksCache = SupplementalBibleManager.IndexStrongDictionary(OneNoteApp, selectedModuleInfo, Logger);
                }

                MainForm.PrepareForExternalProcessing(chaptersCount, 1, BibleCommon.Resources.Constants.LinkSupplementalBibleStart);
                Logger.Preffix = string.Format("{0} {1}/{1}: {2}: ", BibleCommon.Resources.Constants.Stage, stagesCount, BibleCommon.Resources.Constants.LinkSupplementalBible);
                BibleCommon.Services.Logger.LogMessage(Logger.Preffix);
                result = SupplementalBibleManager.LinkSupplementalBibleWithMainBible(OneNoteApp, 0, strongTermLinksCache, Logger);

                MainForm.ExternalProcessingDone(BibleCommon.Resources.Constants.CreateSupplementalBibleFinish);
            }

            return result.Errors.ConvertAll(ex => ex.Message);
        }

        protected override bool CanModuleBeDeleted(int index)
        {
            return index == GetSupplementalModulesCount() - 1;
        }

        protected override void DeleteModule(string moduleShortName)
        {
            int chaptersCount = ModulesManager.GetBibleChaptersCount(moduleShortName);
            MainForm.PrepareForExternalProcessing(chaptersCount, 1, BibleCommon.Resources.Constants.RemoveParallelBibleTranslation);
            var removeResult = SupplementalBibleManager.RemoveLastSupplementalBibleModule(OneNoteApp, Logger);
            MainForm.ExternalProcessingDone(
                removeResult == SupplementalBibleManager.RemoveResult.RemoveLastModule
                    ? BibleCommon.Resources.Constants.RemoveParallelBibleTranslationFinishMessage
                    : BibleCommon.Resources.Constants.RemoveSupplementalBibleFinishMessage);
        }

        protected override string CloseSupplementalNotebookConfirmText
        {
            get { return BibleCommon.Resources.Constants.DeleteSupplementalBibleQuestion; }
        }

        protected override void CloseSupplementalNotebook()
        {
            SupplementalBibleManager.CloseSupplementalBible(OneNoteApp);
        }

        protected override bool IsModuleSupported(ModuleInfo moduleInfo)
        {
            return BibleParallelTranslationManager.IsModuleSupported(moduleInfo);
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
    }
}
