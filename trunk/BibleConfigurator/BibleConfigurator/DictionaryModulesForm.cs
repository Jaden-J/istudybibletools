using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BibleCommon.Services;
using BibleCommon.Common;

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

        protected override void ClearSupplementalModulesInSettingsStorage()
        {
            SettingsManager.Instance.DictionariesModules.Clear();
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
  - необходимо, чтобы в программу были загружены модули типа 'Словарь'  
";
            }
        }

        protected override List<string> CommitChanges(BibleCommon.Common.ModuleInfo selectedModuleInfo)
        {
            DictionaryManager.AddDictionary(OneNoteApp, selectedModuleInfo.ShortName, FolderBrowserDialog.SelectedPath);
            return new List<string>();
        }

        protected override string GetSupplementalModuleName(int index)
        {
            return SettingsManager.Instance.DictionariesModules[index].ModuleName;
        }

        protected override bool CanModuleBeDeleted(int index)
        {
            return true;
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
                && moduleInfo.Type == ModuleType.Dictionary;
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
    }
}
