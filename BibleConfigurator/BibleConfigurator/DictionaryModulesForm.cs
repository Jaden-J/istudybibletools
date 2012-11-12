using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BibleCommon.Services;
using BibleCommon.Common;
using System.Xml;
using BibleCommon.Helpers;
using Microsoft.Office.Interop.OneNote;
using System.Xml.XPath;
using System.Xml.Linq;

namespace BibleConfigurator
{
    public class DictionaryModulesForm: BaseSupplementalForm
    {
        public DictionaryModulesForm(Microsoft.Office.Interop.OneNote.Application oneNoteApp, MainForm form)
            : base(oneNoteApp, form)
        { }

        protected override string GetValidSupplementalNotebookId()
        {
            return SettingsManager.Instance.GetValidDictionariesNotebookId(OneNoteApp, true);
        }        

        protected override int GetSupplementalModulesCount()
        {
            return SettingsManager.Instance.DictionariesModules.Count;
        }

        protected override bool SupplementalModuleAlreadyAdded(string moduleShortName)
        {
            return SettingsManager.Instance.DictionariesModules.Any(m => m.ModuleName == moduleShortName);
        }

        protected override string FormDescription
        {
            get
            {
                return
@"Данная форма предназначена для управления словарями. 

Обратите внимание:  
  - необходимо, чтобы в программу были загружены модули типа 'Словарь';
  - добавление нового словаря может занять несколько минут.
";
            }
        }

        protected override List<string> CommitChanges(BibleCommon.Common.ModuleInfo selectedModuleInfo)
        {
            MainForm.PrepareForExternalProcessing(selectedModuleInfo.NotebooksStructure.DictionaryTermsCount.Value, 1, BibleCommon.Resources.Constants.AddDictionaryStart);
            DictionaryManager.AddDictionary(OneNoteApp, selectedModuleInfo, FolderBrowserDialog.SelectedPath, true);
            Logger.Preffix = string.Format("{0}: ", BibleCommon.Resources.Constants.IndexDictionary);
            DictionaryTermsCacheManager.GenerateCache(OneNoteApp, selectedModuleInfo, Logger);
            MainForm.ExternalProcessingDone(BibleCommon.Resources.Constants.AddDictionaryFinishMessage);

            return new List<string>();
        }

        protected override string GetSupplementalModuleName(int index)
        {
            return SettingsManager.Instance.DictionariesModules[index].ModuleName;
        }

        protected override bool CanModuleBeDeleted(ModuleInfo moduleInfo, int index)
        {
            return moduleInfo.Type == ModuleType.Dictionary
                || (moduleInfo.Type == ModuleType.Strong && !SettingsManager.Instance.SupplementalBibleModules.Any(
                    m => 
                    {
                        var sm = Modules.FirstOrDefault(module => module.ShortName == m.ModuleName);
                        if (sm != null)
                            return sm.Type == ModuleType.Strong;
                        return false;
                    }));
        }

        protected override void DeleteModule(string moduleShortName)
        {
            DictionaryManager.RemoveDictionary(OneNoteApp, moduleShortName);            
        }

        protected override string CloseSupplementalNotebookQuestionText
        {
            get { return BibleCommon.Resources.Constants.DeleteDictionariesNotebookQuestion; }
        }

        protected override void CloseSupplementalNotebook()
        {
            DictionaryManager.CloseDictionariesNotebook(OneNoteApp);
        }

        protected override bool IsModuleSupported(BibleCommon.Common.ModuleInfo moduleInfo)
        {
            return BibleParallelTranslationManager.IsModuleSupported(moduleInfo) 
                && (moduleInfo.Type == ModuleType.Dictionary || moduleInfo.Type == ModuleType.Strong);
        }

        protected override string GetFormText()
        {
            return BibleCommon.Resources.Constants.DictionariesManagement;
        }

        protected override string GetChkUseText()
        {
            return BibleCommon.Resources.Constants.UseDictionaries;
        }

        protected override bool IsBaseModuleSupported()
        {
            return true;
        }

        protected override string DeleteModuleQuestionText
        {
            get { return BibleCommon.Resources.Constants.DeleteThisModuleFromDictionariesNotebookQuestion; }
        }

        protected override bool CanModuleBeAdded(ModuleInfo moduleInfo)
        {
            return moduleInfo.Type == ModuleType.Dictionary;
        }

        protected override bool CanNotebookBeClosed()
        {
            return !SettingsManager.Instance.DictionariesModules.Any(dm => Modules.First(m => m.ShortName == dm.ModuleName).Type == ModuleType.Strong);
        }

        protected override string NotebookCannotBeClosedText
        {
            get { return BibleCommon.Resources.Constants.DictionaryNotebookCannotBeClosed; }
        }       

        protected override string EmbeddedModulesKey
        {
            get { return BibleCommon.Consts.Constants.EmbeddedDictionariesKey; }
        }

        protected override string NotebookIsNotSupplementalBibleMessage
        {
            get { return BibleCommon.Resources.Constants.NotebookIsNotDictionariesNotebook; }
        }

        protected override string SupplementalNotebookWasAddedMessage
        {
            get { return BibleCommon.Resources.Constants.DictionariesNotebookWasAdded; }
        }

        protected override void SaveSupplementalNotebookSettings(string notebookId)
        {
            SettingsManager.Instance.NotebookId_Dictionaries = notebookId;
            SettingsManager.Instance.Save();
        }

        protected override List<string> SaveEmbeddedModuleSettings(EmbeddedModuleInfo embeddedModuleInfo, ModuleInfo moduleInfo, System.Xml.Linq.XElement pageEl)
        {
            var result = new List<string>();

            XElement sectionEl;

            if (string.IsNullOrEmpty(moduleInfo.NotebooksStructure.DictionarySectionGroupName))
                sectionEl = pageEl.Parent;
            else
                sectionEl = pageEl.Parent.Parent;

            var sectionId = (string)sectionEl.Attribute("ID");

            SettingsManager.Instance.DictionariesModules.Add(new StoredModuleInfo(embeddedModuleInfo.ModuleName, embeddedModuleInfo.ModuleVersion, sectionId));

            if (moduleInfo.Type == ModuleType.Strong)
            {
                if (!SettingsManager.Instance.SupplementalBibleModules.Any(m => m.ModuleName == embeddedModuleInfo.ModuleName))
                    result.Add(BibleCommon.Resources.Constants.NeedToAddSupplementalNotebookWithStrongsNumber);
            }

            if (!DictionaryTermsCacheManager.CacheIsActive(moduleInfo.ShortName))
            {
                MainForm.PrepareForExternalProcessing(moduleInfo.NotebooksStructure.DictionaryTermsCount.Value, 1, BibleCommon.Resources.Constants.AddDictionaryStart);
                Logger.Preffix = string.Format("{0} {1}: ", BibleCommon.Resources.Constants.IndexDictionary, moduleInfo.ShortName);
                DictionaryTermsCacheManager.GenerateCache(OneNoteApp, moduleInfo, Logger);
            }

            return result;
        }

        protected override void ClearSupplementalModules()
        {
            SettingsManager.Instance.DictionariesModules.Clear();
        }
    }
}
