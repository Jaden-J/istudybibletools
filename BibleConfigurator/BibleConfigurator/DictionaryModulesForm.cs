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
using BibleCommon.UI.Forms;

namespace BibleConfigurator
{
    public class DictionaryModulesForm: BaseSupplementalForm
    {
        public DictionaryModulesForm(Microsoft.Office.Interop.OneNote.Application oneNoteApp, MainForm form)
            : base(oneNoteApp, form)
        { }

        protected override string GetValidSupplementalNotebookId()
        {
            return SettingsManager.Instance.GetValidDictionariesNotebookId(ref _oneNoteApp, true);
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
                return BibleCommon.Resources.Constants.DictionaryFormDescription;
            }
        }

        protected override ErrorsList CommitChanges(BibleCommon.Common.ModuleInfo selectedModuleInfo)
        {
            if (!string.IsNullOrEmpty(SettingsManager.Instance.GetValidDictionariesNotebookId(ref _oneNoteApp, true)))
                if (!OneNoteUtils.IsNotebookLocal(ref _oneNoteApp, SettingsManager.Instance.NotebookId_Dictionaries))
                    throw new InvalidNotebookException(BibleCommon.Resources.Constants.NotebookIsLocalAndNotSupportedForDictionaries);

            MainForm.PrepareForLongProcessing(selectedModuleInfo.NotebooksStructure.DictionaryTermsCount.Value, 1, BibleCommon.Resources.Constants.AddDictionaryStart);
            DictionaryManager.AddDictionary(ref _oneNoteApp, selectedModuleInfo, FolderBrowserDialog.SelectedPath, true, () => Logger.AbortedByUsers);
            Logger.Preffix = string.Format("{0}: ", BibleCommon.Resources.Constants.IndexDictionary);

            List<string> notFoundTerms;
            DictionaryTermsCacheManager.GenerateCache(ref _oneNoteApp, selectedModuleInfo, Logger, out notFoundTerms);
            MainForm.LongProcessingDone(BibleCommon.Resources.Constants.AddDictionaryFinishMessage);

            if (notFoundTerms != null && notFoundTerms.Count > 0)
            {
                return new ErrorsList(notFoundTerms)
                {
                    ErrorsDecription = BibleCommon.Resources.Constants.DictionaryTermsNotFound
                };
            }
            else
                return null;
        }

        protected override string GetSupplementalModuleName(int index)
        {
            return SettingsManager.Instance.DictionariesModules[index].ModuleName;
        }

        protected override bool CanModuleBeDeleted(ModuleInfo moduleInfo, int index)
        {
            return moduleInfo.Type == ModuleType.Dictionary
                   || (moduleInfo.Type == ModuleType.Strong
                      && !SettingsManager.Instance.SupplementalBibleModules.Any(m => m.ModuleName == moduleInfo.ShortName));
        }

        protected override void DeleteModule(string moduleShortName)
        {
            DictionaryManager.RemoveDictionary(ref _oneNoteApp, moduleShortName, true);
        }

        protected override string CloseSupplementalNotebookQuestionText
        {
            get { return BibleCommon.Resources.Constants.DeleteDictionariesNotebookQuestion; }
        }

        protected override void CloseSupplementalNotebook()
        {
            DictionaryManager.CloseDictionariesNotebook(ref _oneNoteApp, true);
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
            return !(SettingsManager.Instance.DictionariesModules.Any(dm => DictionaryModules[dm.ModuleName].Type == ModuleType.Strong)
                    && SettingsManager.Instance.SupplementalBibleModules.Any(sm => DictionaryModules[sm.ModuleName].Type == ModuleType.Strong)
                    && !string.IsNullOrEmpty(SettingsManager.Instance.GetValidSupplementalBibleNotebookId(ref _oneNoteApp)));
        }

        protected override string NotebookCannotBeClosedText
        {
            get { return BibleCommon.Resources.Constants.DictionariesNotebookCannotBeClosed; }
        }       

        protected override string EmbeddedModulesKey
        {
            get { return BibleCommon.Consts.Constants.Key_EmbeddedDictionaries; }
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
                MainForm.PrepareForLongProcessing(moduleInfo.NotebooksStructure.DictionaryTermsCount.Value, 1, BibleCommon.Resources.Constants.AddDictionaryStart);
                Logger.Preffix = string.Format("{0} {1}: ", BibleCommon.Resources.Constants.IndexDictionary, moduleInfo.ShortName);

                List<string> notFoundTerms;
                DictionaryTermsCacheManager.GenerateCache(ref _oneNoteApp, moduleInfo, Logger, out notFoundTerms);

                if (notFoundTerms != null && notFoundTerms.Count > 0)
                {
                    using (var form = new ErrorsForm())
                    {
                        form.AllErrors.Add(new ErrorsList(notFoundTerms)
                        {
                            ErrorsDecription = BibleCommon.Resources.Constants.DictionaryTermsNotFound
                        });
                        form.ShowDialog();
                    }
                }
            }

            return result;
        }

        protected override void ClearSupplementalModules()
        {
            SettingsManager.Instance.DictionariesModules.Clear();
        }

        protected override bool AreThereModulesToAdd()
        {
            return Modules.Any(m => m.Type == ModuleType.Dictionary && !SupplementalModuleAlreadyAdded(m.ShortName));            
        }

        protected override string GetPostCommitErrorMessage(ModuleInfo selectedModuleInfo)
        {
            return null;
        }

        protected override void CheckIfExistingNotebookCanBeUsed(string notebookId)
        {
            if (!OneNoteUtils.IsNotebookLocal(ref _oneNoteApp, SettingsManager.Instance.NotebookId_Dictionaries))            
                throw new InvalidNotebookException(BibleCommon.Resources.Constants.NotebookIsLocalAndNotSupportedForDictionaries);            
        }
    }
}
