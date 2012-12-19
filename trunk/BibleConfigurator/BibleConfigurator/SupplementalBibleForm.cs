using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BibleCommon.Services;
using BibleCommon.Common;
using System.Xml;
using BibleCommon.Helpers;
using Microsoft.Office.Interop.OneNote;
using System.Xml.Linq;
using System.Xml.XPath;
using BibleCommon.UI.Forms;
using System.Windows.Forms;
using System.IO;

namespace BibleConfigurator
{
    public class SupplementalBibleForm: BaseSupplementalForm
    {
        public SupplementalBibleForm(Microsoft.Office.Interop.OneNote.Application oneNoteApp, MainForm form)
            : base(oneNoteApp, form)
        { }        

        protected override string GetValidSupplementalNotebookId()
        {
            return SettingsManager.Instance.GetValidSupplementalBibleNotebookId(ref _oneNoteApp, true);
        }        

        protected override int GetSupplementalModulesCount()
        {
            return SettingsManager.Instance.SupplementalBibleModules.Count;
        }

        protected override string GetSupplementalModuleName(int index)
        {
            return SettingsManager.Instance.SupplementalBibleModules[index].ModuleName;
        }

        protected override string FormDescription
        {
            get
            {
                return BibleCommon.Resources.Constants.SupplementalBibleFormDescription;
            }
        }

        protected override ErrorsList CommitChanges(ModuleInfo selectedModuleInfo)
        {
            BibleParallelTranslationConnectionResult result;

            Dictionary<string, string> strongTermLinksCache = null;

            if (SettingsManager.Instance.SupplementalBibleModules.Count > 0)
            {
                if (!CheckModules(DictionaryModules[SettingsManager.Instance.SupplementalBibleModules.First().ModuleName],
                    selectedModuleInfo, BibleCommon.Resources.Constants.ContinueAddingParallelBible))
                    return null;

                int stagesCount = selectedModuleInfo.Type == ModuleType.Strong ? 2 : 1;

                int chaptersCount = ModulesManager.GetBibleChaptersCount(SettingsManager.Instance.SupplementalBibleModules.First().ModuleName, false);                

                if (selectedModuleInfo.Type == ModuleType.Strong)
                {
                    MainForm.PrepareForLongProcessing(1, 1, BibleCommon.Resources.Constants.AddParallelBibleTranslationStart);
                    DictionaryManager.AddDictionary(ref _oneNoteApp, selectedModuleInfo, FolderBrowserDialog.SelectedPath, true, () => Logger.AbortedByUsers);
                    strongTermLinksCache = RunIndexStrongStage(selectedModuleInfo, 1, stagesCount, false);
                    MainForm.PrepareForLongProcessing(chaptersCount, 1, string.Empty);                
                }
                else
                    MainForm.PrepareForLongProcessing(chaptersCount, 1, BibleCommon.Resources.Constants.AddParallelBibleTranslationStart);                

                string stagesString = stagesCount == 1 ? string.Empty : string.Format("{0} {1}/{1}: ", BibleCommon.Resources.Constants.Stage, stagesCount);
                Logger.Preffix = string.Format("{0}{1}: ", stagesString, BibleCommon.Resources.Constants.AddParallelBibleTranslation); 
                BibleCommon.Services.Logger.LogMessageParams(Logger.Preffix);
                result = SupplementalBibleManager.AddParallelBible(ref _oneNoteApp, selectedModuleInfo, strongTermLinksCache, Logger);

                MainForm.LongProcessingDone(BibleCommon.Resources.Constants.AddParallelBibleTranslationFinishMessage);
            }
            else
            {
                if (!CheckModules(selectedModuleInfo, SettingsManager.Instance.CurrentModuleCached, BibleCommon.Resources.Constants.ContinueCreatingSupplementalBible))
                    return null;

                int stagesCount = selectedModuleInfo.Type == ModuleType.Strong ? 3 : 2;

                int chaptersCount = ModulesManager.GetBibleChaptersCount(selectedModuleInfo.ShortName, false);
                MainForm.PrepareForLongProcessing(chaptersCount, 1, BibleCommon.Resources.Constants.CreateSupplementalBibleStart);
                Logger.Preffix = string.Format("{0} 1/{1}: {2}: ", BibleCommon.Resources.Constants.Stage, stagesCount, BibleCommon.Resources.Constants.CreateSupplementalBible);
                BibleCommon.Services.Logger.LogMessageParams(Logger.Preffix);

                if (selectedModuleInfo.Type == ModuleType.Strong)
                    DictionaryManager.AddDictionary(ref _oneNoteApp, selectedModuleInfo, FolderBrowserDialog.SelectedPath, false, () => Logger.AbortedByUsers);                
                SupplementalBibleManager.CreateSupplementalBible(ref _oneNoteApp, selectedModuleInfo, FolderBrowserDialog.SelectedPath, Logger);
                
                if (selectedModuleInfo.Type == ModuleType.Strong)                
                    strongTermLinksCache = RunIndexStrongStage(selectedModuleInfo, 2, stagesCount, true);                                    

                MainForm.PrepareForLongProcessing(chaptersCount, 1, BibleCommon.Resources.Constants.LinkSupplementalBibleStart);
                Logger.Preffix = string.Format("{0} {1}/{1}: {2}: ", BibleCommon.Resources.Constants.Stage, stagesCount, BibleCommon.Resources.Constants.LinkSupplementalBible);
                BibleCommon.Services.Logger.LogMessageParams(Logger.Preffix);
                result = SupplementalBibleManager.LinkSupplementalBibleWithPrimaryBible(ref _oneNoteApp, strongTermLinksCache, Logger);

                MainForm.LongProcessingDone(BibleCommon.Resources.Constants.CreateSupplementalBibleFinish);
            }

            return new ErrorsList(result.Errors.ConvertAll(ex => ex.Message));
        }

        private bool CheckModules(ModuleInfo baseModule, ModuleInfo parallelModule, string messageToContinue)
        {
            var notFoundVerses = new List<ErrorsList>();

            MainForm.PrepareForLongProcessing(SettingsManager.Instance.SupplementalBibleModules.Count > 0 ? 3 : 2, 1, BibleCommon.Resources.Constants.ParallelModuleChecking);                        
            
            var errors = ParallelBibleCheckerForm.CheckModule(baseModule.ShortName, parallelModule.ShortName);
            MainForm.PerformProgressStep(BibleCommon.Resources.Constants.ParallelModuleChecking);

            if (SettingsManager.Instance.SupplementalBibleModules.Count > 0)
            {
                notFoundVerses.Add(new ErrorsList(BibleParallelTranslationManager.CheckForInconsistencies(parallelModule.ShortName, baseModule.ShortName))
                                   {
                                       ErrorsDecription = BibleCommon.Resources.Constants.NotExistingInFirstSupplemenBibleVersesFound
                                   });
                MainForm.PerformProgressStep(BibleCommon.Resources.Constants.ParallelModuleChecking);

                notFoundVerses.Add(new ErrorsList(BibleParallelTranslationManager.CheckForInconsistencies(parallelModule.ShortName, SettingsManager.Instance.ModuleShortName))
                {
                    ErrorsDecription = BibleCommon.Resources.Constants.NotExistingInPrimaryBibleVersesFound
                });
                MainForm.PerformProgressStep(BibleCommon.Resources.Constants.ParallelModuleCheckFinish);
            }
            else
            {
                notFoundVerses.Add(new ErrorsList(BibleParallelTranslationManager.CheckForInconsistencies(baseModule.ShortName, SettingsManager.Instance.ModuleShortName))
                {
                    ErrorsDecription = BibleCommon.Resources.Constants.NotExistingInPrimaryBibleVersesFound
                });
                MainForm.PerformProgressStep(BibleCommon.Resources.Constants.ParallelModuleCheckFinish);
            }            

            var needToContinue = ShowWarnings(baseModule, parallelModule, notFoundVerses, messageToContinue);

            if (needToContinue)
                needToContinue = ShowErrors(baseModule, parallelModule, errors, messageToContinue);

            return needToContinue;
        }

        private bool ShowErrors(ModuleInfo baseModule, ModuleInfo parallelModule, ErrorsList errors, string messageToContinue)
        {
            if (errors != null)
            {
                var errorsFile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), 
                    string.Format("{0}--{1}.errors.txt", baseModule.ShortName, parallelModule.ShortName));

                using (var form = new ErrorsForm())
                {
                    form.AllErrors.Add(errors);
                    form.SaveErrorsToFile(errorsFile);
                }

                using (var form = new MessageForm(string.Format("{0}{3}{1}{3}{2}",
                                                        string.Format(BibleCommon.Resources.Constants.ThereAreErrorsOnModulesMerging,
                                                                        baseModule.ShortName, baseModule.Version,
                                                                        parallelModule.ShortName, parallelModule.Version,
                                                                        BibleCommon.Resources.Constants.WebSiteUrl),
                                                        string.Format(BibleCommon.Resources.Constants.ErrorsAreSavedInFile, errorsFile),
                                                        messageToContinue, Environment.NewLine),
                                                  BibleCommon.Resources.Constants.Warning,
                                                  System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question))
                {
                    if (form.ShowDialog() != System.Windows.Forms.DialogResult.Yes)
                    {
                        MainForm.LongProcessingDone(string.Empty);
                        return false;
                    }
                }
            }

            return true;
        }

        private bool ShowWarnings(ModuleInfo baseModule, ModuleInfo parallelModule, List<ErrorsList> notFoundVerses, string messageToContinue)
        {
            if (notFoundVerses.Any(wrngs => wrngs.Count > 0))
            {
                var notFoundVersesFile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
                    string.Format("{0}--{1}.warnings.txt", baseModule.ShortName, parallelModule.ShortName));

                using (var form = new ErrorsForm())
                {
                    form.AllErrors = notFoundVerses;
                    form.SaveErrorsToFile(notFoundVersesFile);
                }

                var warningMessage = string.Empty;
                foreach (var nfv in notFoundVerses)
                {
                    if (nfv.Count > 0)
                    {
                        if (!string.IsNullOrEmpty(warningMessage))
                            warningMessage += Environment.NewLine;
                        warningMessage += nfv.ErrorsDecription;
                    }
                }

                warningMessage += Environment.NewLine + string.Format(BibleCommon.Resources.Constants.ErrorsAreSavedInFile, notFoundVersesFile);
                warningMessage += Environment.NewLine + messageToContinue;

                using (var form = new MessageForm(warningMessage,
                                                 BibleCommon.Resources.Constants.Warning,
                                                 System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question))
                {
                    if (form.ShowDialog() != System.Windows.Forms.DialogResult.Yes)
                    {
                        MainForm.LongProcessingDone(string.Empty);
                        return false;
                    }
                }
            }

            return true;
        }

        private Dictionary<string, string> RunIndexStrongStage(ModuleInfo moduleInfo, int stage, int stagesCount, bool checkPagesCount)
        {
            int strongTermsCount = moduleInfo.NotebooksStructure.DictionaryTermsCount.GetValueOrDefault(BibleCommon.Consts.Constants.DefaultStrongNumbersCount);
            var pagesCount = moduleInfo.NotebooksStructure.DictionaryPagesCount;
            var dictionaryModuleInfo = SettingsManager.Instance.DictionariesModules.FirstOrDefault(m => m.ModuleName == moduleInfo.ShortName);

            MainForm.PrepareForLongProcessing(strongTermsCount, 1, BibleCommon.Resources.Constants.IndexStrongDictionaryStart);
            Logger.Preffix = string.Format("{0} {1}/{2}: {3}: ", BibleCommon.Resources.Constants.Stage, stage, stagesCount, BibleCommon.Resources.Constants.IndexStrongDictionary);
            BibleCommon.Services.Logger.LogMessageParams(Logger.Preffix);

            if (checkPagesCount)
                DictionaryManager.WaitWhileDictionaryIsCreating(ref _oneNoteApp, dictionaryModuleInfo.SectionId, pagesCount, 0, () => Logger.AbortedByUsers);  // повторный раз проверяем, что все страницы загрузились

            Dictionary<string, string> result;

            if (DictionaryTermsCacheManager.CacheIsActive(moduleInfo.ShortName))
                result = DictionaryTermsCacheManager.LoadCachedDictionary(moduleInfo.ShortName);
            else
            {
                List<string> notFoundTerms;
                result = DictionaryTermsCacheManager.GenerateCache(ref _oneNoteApp, moduleInfo, Logger, out notFoundTerms);

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

        protected override bool CanModuleBeDeleted(ModuleInfo moduleInfo, int index)
        {
            return index != 0 || index == GetSupplementalModulesCount() - 1;
        }

        protected override void DeleteModule(string moduleShortName)
        {
            int chaptersCount = ModulesManager.GetBibleChaptersCount(moduleShortName, false);            
            MainForm.PrepareForLongProcessing(chaptersCount, 1, BibleCommon.Resources.Constants.RemoveParallelBibleTranslationStartMessage);
            Logger.Preffix = string.Format("{0}: ", BibleCommon.Resources.Constants.RemoveParallelBibleTranslation);

            var removeStrongDictionaryFromNotebook = true;            

            if (DictionaryModules[moduleShortName].Type == ModuleType.Strong)
            {
                if (MessageBox.Show(BibleCommon.Resources.Constants.RemoveStrongDictionaryFromNotebookQuestion,
                    BibleCommon.Resources.Constants.Warning, MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1)
                        == System.Windows.Forms.DialogResult.No)
                    removeStrongDictionaryFromNotebook = false;
            }

            var removeResult = SupplementalBibleManager.RemoveSupplementalBibleModule(ref _oneNoteApp, moduleShortName, removeStrongDictionaryFromNotebook, Logger);
            MainForm.LongProcessingDone(
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
            var removeStrongDictionaryFromNotebook = true;

            foreach (var module in SettingsManager.Instance.SupplementalBibleModules)
            {
                if (DictionaryModules[module.ModuleName].Type == ModuleType.Strong)
                {
                    if (MessageBox.Show(BibleCommon.Resources.Constants.RemoveStrongDictionaryFromNotebookQuestion,
                        BibleCommon.Resources.Constants.Warning, MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1)
                            == System.Windows.Forms.DialogResult.No)                    
                        removeStrongDictionaryFromNotebook = false;

                    break;
                }
            }

            SupplementalBibleManager.CloseSupplementalBible(ref _oneNoteApp, removeStrongDictionaryFromNotebook);
        }

        protected override bool IsModuleSupported(ModuleInfo moduleInfo)
        {
            return BibleParallelTranslationManager.IsModuleSupported(moduleInfo) 
                && (moduleInfo.Type == ModuleType.Bible || moduleInfo.Type == ModuleType.Strong);
        }

        protected override bool SupplementalModuleAlreadyAdded(string moduleShortName)
        {
            return SettingsManager.Instance.SupplementalBibleModules.Any(m => m.ModuleName == moduleShortName);
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
            return BibleParallelTranslationManager.IsModuleSupported(SettingsManager.Instance.CurrentModuleCached);
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

        protected override string EmbeddedModulesKey
        {
            get { return BibleCommon.Consts.Constants.Key_EmbeddedSupplementalModules; }
        }

        protected override string NotebookIsNotSupplementalBibleMessage
        {
            get { return BibleCommon.Resources.Constants.NotebookIsNotSupplementalBible; }
        }

        protected override string SupplementalNotebookWasAddedMessage
        {
            get { return BibleCommon.Resources.Constants.SupplementalNotebookWasAdded; }
        }

        protected override void SaveSupplementalNotebookSettings(string notebookId)
        {
            SettingsManager.Instance.NotebookId_SupplementalBible = notebookId;
            SettingsManager.Instance.Save();
        }

        protected override List<string> SaveEmbeddedModuleSettings(EmbeddedModuleInfo embeddedModuleInfo, ModuleInfo moduleInfo, XElement pageEl)
        {
            var result = new List<string>();

            SettingsManager.Instance.SupplementalBibleModules.Add(new StoredModuleInfo(embeddedModuleInfo.ModuleName, embeddedModuleInfo.ModuleVersion));
            if (moduleInfo.Type == ModuleType.Strong)
            {
                if (!SettingsManager.Instance.DictionariesModules.Any(m => m.ModuleName == embeddedModuleInfo.ModuleName))
                    result.Add(BibleCommon.Resources.Constants.NeedToAddDictionaryNotebookWithStrongsNumber);
            }

            return result;
        }

        protected override void ClearSupplementalModules()
        {
            SettingsManager.Instance.SupplementalBibleModules.Clear();
        }

        protected override bool AreThereModulesToAdd()
        {
            return Modules.Any(m => IsModuleSupported(m) && !SupplementalModuleAlreadyAdded(m.ShortName));            
        }

        protected override string GetPostCommitErrorMessage(ModuleInfo selectedModuleInfo)
        {
            var primaryModule = DictionaryModules[SettingsManager.Instance.SupplementalBibleModules.First().ModuleName];
            var parallelModule = SettingsManager.Instance.SupplementalBibleModules.Count > 1
                                    ? selectedModuleInfo
                                    : SettingsManager.Instance.CurrentModuleCached;

            return string.Format("{0} {1}",
                string.Format(BibleCommon.Resources.Constants.ThereAreErrorsOnModulesMerging,
                                        primaryModule.ShortName, primaryModule.Version,
                                        parallelModule.ShortName, parallelModule.Version,
                                        BibleCommon.Resources.Constants.WebSiteUrl),
                BibleCommon.Resources.Constants.ThereAreErrorsAfterParallelModuleWasAdded);
        }

        protected override void CheckIfExistingNotebookCanBeUsed(string notebookId)
        {
            // поддерживаем как локальные, так и облачные записные книжки
        }
    }
}
