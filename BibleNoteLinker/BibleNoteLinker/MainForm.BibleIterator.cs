using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.OneNote;
using BibleCommon.Services;
using System.Xml.Linq;
using BibleCommon.Consts;
using BibleCommon.Helpers;
using BibleCommon.Common;
using System.Xml.XPath;

namespace BibleNoteLinker
{
    public partial class MainForm
    {
        
        private const int ApproximatePageVersesCount = 100;
        private int _pagesForAnalyzeCount;

        protected int StagesCount { get; set; }

        private void StartAnalyze()
        {            
            pbMain.Value = 0;

            try
            {
                OneNoteLocker.UnlockBible(ref _oneNoteApp, true, () => _processAbortedByUser);
            }
            catch (NotSupportedException)
            {
                //todo: log it
            }

            Exception getCurrentPageException = null;
            NotebookIterator.PageInfo currentPage = null;
            try
            {
                currentPage = OneNoteUtils.GetCurrentPageInfo(ref _oneNoteApp);
            }
            catch (ProgramException ex)
            {
                getCurrentPageException = ex;
            }

            StagesCount = GetStagesCount();

            if (!rbAnalyzeCurrentPage.Checked)
            {
                List<NotebookIterator.NotebookInfo> notebooks = GetNotebooksInfo();
                _pagesForAnalyzeCount = notebooks.Sum(notebook => notebook.PagesCount);

                pbMain.Maximum = _pagesForAnalyzeCount > 1 ? _pagesForAnalyzeCount : ApproximatePageVersesCount;

                pbMain.PerformStep();
                Logger.LogMessageParams(Helper.GetRightFoundPagesString(_pagesForAnalyzeCount));

                foreach (NotebookIterator.NotebookInfo notebook in notebooks)
                    ProcessNotebook(notebook);
            }
            else
            {
                if (getCurrentPageException != null)
                    throw getCurrentPageException;

                _pagesForAnalyzeCount = 1;
                string message = BibleCommon.Resources.Constants.ProcessCurrentPage;

                pbMain.Maximum = ApproximatePageVersesCount;

                LogHighLevelMessage(message, 1, StagesCount);
                Logger.LogMessageParams(message);
                Logger.MoveLevel(1);
                ProcessPage(currentPage, null);
                Logger.MoveLevel(-1);
            }

            if (_pagesForAnalyzeCount > 0)
            {
                var currentStep = 2;
                if (!SettingsManager.Instance.StoreNotesPagesInFolder)
                {
                    CommitPagesInOneNote(BibleCommon.Resources.Constants.NoteLinkerNotesPagesUpdating, currentStep++, ApplicationCache.PageType.NotesPage);

                    SyncNotesPagesContainer();   // эта задача асинхронная, поэтому не выделаем как отдельный этап

                    SortNotesPages();  // это происходит очень быстро, поэтому не выделяем как отдельный этап                

                    CommitNotesPagesHierarchy(currentStep++);
                }

                if (SettingsManager.Instance.StoreNotesPagesInFolder)
                    CommitNotesPagesInFileSystem(currentStep++);

                if (!SettingsManager.Instance.IsInIntegratedMode)
                    UpdateLinksToNotesPages(currentStep++);                

                if (!SettingsManager.Instance.IsInIntegratedMode)
                    CommitPagesInOneNote(BibleCommon.Resources.Constants.NoteLinkerBiblePagesUpdating, currentStep++, null);                
            }

            if (SettingsManager.Instance.StoreNotesPagesInFolder && chkForce.Checked && rbAnalyzeAllPages.Checked)
            {
                NotesPageManagerFS.UpdateResources();
            }

            OneNoteUtils.UseOneNoteAPI(ref _oneNoteApp, () =>
            {
                if (_oneNoteApp.Windows.CurrentWindow != null && currentPage != null)
                {
                    _oneNoteApp.NavigateTo(currentPage.Id, null);                    
                }
            });

            //OneNoteUtils.SetActiveCurrentWindow(ref _oneNoteApp);
        }

        private int GetStagesCount()
        {
            if (SettingsManager.Instance.StoreNotesPagesInFolder)
            {
                if (SettingsManager.Instance.IsInIntegratedMode)
                    return 2;
                else
                    return 4;
            }
            else
                return 5;            
        }

        private void CommitNotesPagesInFileSystem(int stage)
        {
            string message = BibleCommon.Resources.Constants.NoteLinkerNotesPagesUpdating;
            LogHighLevelMessage(message, stage, StagesCount);
            int allPagesCount = ApplicationCache.Instance.NotesPageDataList.Count;
            Logger.LogMessageParams(string.Format("{0} ({1})",
                message, Helper.GetRightPagesString(allPagesCount)));
            pbMain.Maximum = allPagesCount;
            pbMain.Value = 0;
            pbMain.PerformStep();

            int processedPagesCount = 0;

            for (var i = 0; i < ApplicationCache.Instance.NotesPageDataList.Count; i++)
            {
                LogHighLevelAdditionalMessage(string.Format(": {0}/{1}", ++processedPagesCount, allPagesCount));
                
                try
                {
                    ApplicationCache.Instance.NotesPageDataList[i].Serialize(ref _oneNoteApp);
                    ApplicationCache.Instance.NotesPageDataList[i] = null;  // освобождаем память. Так как таких объектов много, а память ещё нужна для обновления страниц в OneNote.
                }
                catch (Exception ex)
                {
                    Logger.LogError(string.Format(BibleCommon.Resources.Constants.ErrorWhilePageProcessing, ApplicationCache.Instance.NotesPageDataList[i].PageName), ex);
                }

                PerformProcessStep();
            }
        }

        private void SyncNotesPagesContainer()
        {
            OneNoteUtils.UseOneNoteAPI(ref _oneNoteApp, () =>
            {
                _oneNoteApp.SyncHierarchy(!string.IsNullOrEmpty(SettingsManager.Instance.SectionGroupId_BibleNotesPages)
                                          ? SettingsManager.Instance.SectionGroupId_BibleNotesPages
                                          : SettingsManager.Instance.NotebookId_BibleNotesPages);
            });
        }

        private void PerformProcessStep()
        {
            System.Windows.Forms.Application.DoEvents();
            if (_processAbortedByUser)
                throw new ProcessAbortedByUserException();
            pbMain.PerformStep();
        }

        private void CommitNotesPagesHierarchy(int stage)
        {
            string message = BibleCommon.Resources.Constants.NoteLilnkerHierarchyUpdating;
            LogHighLevelMessage(message, stage, StagesCount);
            int allPagesCount = 0;
            int processedPagesCount = 0;
            Logger.LogMessage(message, true, false);
            ApplicationCache.Instance.CommitAllModifiedHierarchy(ref _oneNoteApp,
                pagesCount =>
                {
                    allPagesCount = pagesCount;
                    pbMain.Maximum = pagesCount;
                    pbMain.Value = 0;                    
                    Logger.LogMessage(string.Format(" ({0})", Helper.GetRightPagesString(pagesCount)), false, true, false);
                    //pbMain.PerformStep();
                    //LogHighLevelAdditionalMessage(string.Format(": {0}/{1}", ++processedPagesCount, allPagesCount));
                },
                pageContent => 
                {
                    PerformProcessStep();
                    LogHighLevelAdditionalMessage(string.Format(": {0}/{1}", ++processedPagesCount, allPagesCount));
                });
        }

        private void SortNotesPages()
        {
            //Сортировка страниц 'Сводные заметок'
            foreach (var sortPageInfo in ApplicationCache.Instance.SortVerseLinkPagesInfo)
            {
                try
                {
                    VerseLinkManager.SortVerseLinkPages(ref _oneNoteApp,
                        sortPageInfo.SectionId, sortPageInfo.PageId, sortPageInfo.ParentPageId, sortPageInfo.PageLevel);
                }
                catch (Exception ex)
                {
                    Logger.LogError(string.Format("{0} '{1}'", BibleCommon.Resources.Constants.NoteLinkerErrorWhilePageSorting, sortPageInfo.PageId), ex);
                }
            }
        }

        private void UpdateLinksToNotesPages(int stage)
        {
            string message = BibleCommon.Resources.Constants.NoteLinkerLinksToNotesPagesUpdating;
            LogHighLevelMessage(message, stage, StagesCount);
            int allPagesCount = ApplicationCache.Instance.BiblePagesWithUpdatedLinksToNotesPages.Values.Count;
            Logger.LogMessageParams(string.Format("{0} ({1})",
                message, Helper.GetRightPagesString(allPagesCount)));
            pbMain.Maximum = allPagesCount;
            pbMain.Value = 0;
            pbMain.PerformStep();

            int processedPagesCount = 0;
            var relinkNotesManager = new RelinkAllBibleNotesManager();
            var locale = LanguageManager.GetCurrentCultureInfoBaseLocale();
            foreach (var processedBiblePageId in ApplicationCache.Instance.BiblePagesWithUpdatedLinksToNotesPages.Values)
            {
                LogHighLevelAdditionalMessage(string.Format(": {0}/{1}", ++processedPagesCount, allPagesCount));

                try
                {
                    var vp = processedBiblePageId.ChapterPointer;
                    if (string.IsNullOrEmpty(processedBiblePageId.PageId))
                    {                        
                        var hierarchySearchResult = HierarchySearchManager.GetHierarchyObject(ref _oneNoteApp, SettingsManager.Instance.NotebookId_Bible, ref vp, HierarchySearchManager.FindVerseLevel.OnlyFirstVerse, null, null);
                        if (hierarchySearchResult.FoundSuccessfully)
                        {
                            processedBiblePageId.SectionId = hierarchySearchResult.HierarchyObjectInfo.SectionId;
                            processedBiblePageId.PageId = hierarchySearchResult.HierarchyObjectInfo.PageId;
                            processedBiblePageId.PageName = hierarchySearchResult.HierarchyObjectInfo.PageName;
                            processedBiblePageId.LoadedFromCache = hierarchySearchResult.HierarchyObjectInfo.LoadedFromCache;
                        }
                    }

                    if (!string.IsNullOrEmpty(processedBiblePageId.PageId))
                    {
                        var processedBiblePageIdLocal = processedBiblePageId;
                        HierarchySearchManager.UseHierarchyObjectSafe(ref _oneNoteApp, ref processedBiblePageIdLocal, ref vp, (verseHierarchyInfoSafe) =>
                        {
                            relinkNotesManager.RelinkBiblePageNotes(ref _oneNoteApp, verseHierarchyInfoSafe.SectionId, verseHierarchyInfoSafe.PageId,
                                                        verseHierarchyInfoSafe.PageName, vp, locale);
                            return true;
                        }, null, null);
                    }
                }
                catch (Exception ex)
                {
                    Logger.LogError(string.Format(BibleCommon.Resources.Constants.ErrorWhilePageProcessing, processedBiblePageId.PageName), ex);
                }

                PerformProcessStep();
            }
        }

        private void CommitPagesInOneNote(string startMessage, int stage, ApplicationCache.PageType? pagesType)
        {   
            LogHighLevelMessage(startMessage, stage, StagesCount);
            Logger.LogMessage(startMessage, true, false);
            int allPagesCount = 0;
            int processedPagesCount = 0;
            //Logger.LogMessage(startMessage, true, false);
            ApplicationCache.Instance.CommitAllModifiedPages(ref _oneNoteApp, false,
                pageContent => pagesType.HasValue ? pageContent.PageType == pagesType : true,
                pagesCount =>
                {
                    allPagesCount = pagesCount;
                    pbMain.Maximum = pagesCount;
                    pbMain.Value = 0;                    
                    Logger.LogMessage(string.Format(" ({0})", Helper.GetRightPagesString(pagesCount)), false, true, false);
                    //pbMain.PerformStep();
                    //LogHighLevelAdditionalMessage(string.Format(": {0}/{1}", ++processedPagesCount, allPagesCount));
                },
                pageContent =>
                {
                    PerformProcessStep();
                    LogHighLevelAdditionalMessage(string.Format(": {0}/{1}", ++processedPagesCount, allPagesCount));
                });
        }            

        public void ProcessNotebook(NotebookIterator.NotebookInfo notebook)
        {
            if (notebook.PagesCount > 0)
            {
                Logger.LogMessageParams("{0}: '{1}'", BibleCommon.Resources.Constants.NoteLinkerProcessNotebook, notebook.Title);
                Logger.MoveLevel(1);
                
                ProcessSectionGroup(notebook.RootSectionGroup, true, null);

                Logger.MoveLevel(-1);
            }
        }

        private void LogHighLevelAdditionalMessage(string message)
        {
            lblProgress.Text = _highLevelMessage + message;
        }

        private string _highLevelMessage;
        private void LogHighLevelMessage(string message, int? stage, int? maxStageCount)
        {
            int maxCount = 50;
            if (message.Length > maxCount)
                message = message.Substring(0, maxCount) + "...";

            if (stage.HasValue)
                message = string.Format("{0} {1}/{2}: {3}", BibleCommon.Resources.Constants.Stage, stage, maxStageCount, message);

            _highLevelMessage = message;
            lblProgress.Text = message;
            System.Windows.Forms.Application.DoEvents();
        }

        private void ProcessSectionGroup(BibleCommon.Services.NotebookIterator.SectionGroupInfo sectionGroup, bool isRoot, bool? doNotAnalyze)
        {
            if (!isRoot)
            {
                Logger.LogMessageParams("{0} '{1}'", BibleCommon.Resources.Constants.ProcessSectionGroup, sectionGroup.Title);
                Logger.MoveLevel(1);                
            }

            doNotAnalyze = doNotAnalyze.GetValueOrDefault(false) || StringUtils.IndexOfAny(sectionGroup.Title, Constants.DoNotAnalyzeSymbol1, Constants.DoNotAnalyzeSymbol2) > -1;

            foreach (BibleCommon.Services.NotebookIterator.SectionInfo section in sectionGroup.Sections)
            {
                ProcessSection(section, doNotAnalyze);                
            }

            foreach (BibleCommon.Services.NotebookIterator.SectionGroupInfo subSectionGroup in sectionGroup.SectionGroups)
            {
                ProcessSectionGroup(subSectionGroup, false, doNotAnalyze);
            }

            if (!isRoot)
                Logger.MoveLevel(-1);
        }

        private void ProcessSection(NotebookIterator.SectionInfo section, bool? doNotAnalyze)
        {
            Logger.LogMessageParams("{0} '{1}'", BibleCommon.Resources.Constants.ProcessSection, section.Title);
            Logger.MoveLevel(1);

            doNotAnalyze = doNotAnalyze.GetValueOrDefault(false) || StringUtils.IndexOfAny(section.Title, Constants.DoNotAnalyzeSymbol1, Constants.DoNotAnalyzeSymbol2) > -1; 

            foreach (BibleCommon.Services.NotebookIterator.PageInfo page in section.Pages)
            {
                string message = string.Format("{0} '{1}'", BibleCommon.Resources.Constants.ProcessPage, page.Title);
                LogHighLevelMessage(message, 1, StagesCount);
                Logger.LogMessageParams(message);
                Logger.MoveLevel(1);                

                ProcessPage(page, doNotAnalyze);

                Logger.MoveLevel(-1);
            }

            Logger.MoveLevel(-1);
        }

        private void ProcessPage(NotebookIterator.PageInfo page, bool? doNotAnalyze)
        {
            var noteLinkManager = new NoteLinkManager() { AnalyzeAllPages = rbAnalyzeAllPages.Checked };
            noteLinkManager.OnNextVerseProcess += new EventHandler<NoteLinkManager.ProcessVerseEventArgs>(noteLinkManager_OnNextVerseProcess);
            noteLinkManager.LinkPageVerses(ref _oneNoteApp, page.NotebookId, page.Id, NoteLinkManager.AnalyzeDepth.Full, chkForce.Checked, doNotAnalyze);
            

            pbMain.PerformStep();
            System.Windows.Forms.Application.DoEvents();            

            if (_processAbortedByUser)
                throw new ProcessAbortedByUserException();
        }

        void noteLinkManager_OnNextVerseProcess(object sender, NoteLinkManager.ProcessVerseEventArgs e)
        {
            System.Windows.Forms.Application.DoEvents();
            e.CancelProcess = _processAbortedByUser;                       

            if (e.FoundVerse)
            {
                if (_pagesForAnalyzeCount == 1)
                {
                    if (pbMain.Value == pbMain.Maximum)
                        pbMain.Value = 0;
                    pbMain.PerformStep();
                }

                LogHighLevelAdditionalMessage(string.Format(": {0}", e.VersePointer.OriginalVerseName));
            }
        }

        private List<NotebookIterator.NotebookInfo> GetNotebooksInfo()
        {
            var iterator = new NotebookIterator();
            var result = new List<NotebookIterator.NotebookInfo>();

            Func<NotebookIterator.PageInfo, bool> filter = null;
            if (rbAnalyzeChangedPages.Checked)
                filter = IsPageWasModifiedAfterLastAnalyze;

            if (SettingsManager.Instance.IsSingleNotebook)
            {
                result.Add(iterator.GetNotebookPages(ref _oneNoteApp, SettingsManager.Instance.NotebookId_Bible, SettingsManager.Instance.SectionGroupId_BibleStudy, filter));
                //result.Add(iterator.GetNotebookPages(SettingsManager.Instance.NotebookId_Bible, SettingsManager.Instance.SectionGroupId_BibleComments, filter));
            }
            else
            {
                foreach (var notebookInfo in SettingsManager.Instance.SelectedNotebooksForAnalyze.ToArray())
                {
                    try
                    {
                        result.Add(iterator.GetNotebookPages(ref _oneNoteApp, notebookInfo.NotebookId, null, filter));
                    }
                    catch (Exception ex)
                    {
                        if (OneNoteUtils.IsError(ex, Error.hrObjectDoesNotExist))
                        {
                            SettingsManager.Instance.SelectedNotebooksForAnalyze.Remove(notebookInfo);
                            SettingsManager.Instance.Save();
                        }
                        else
                            throw;
                    }
                }
            }

            return result;
        }

        private bool IsPageWasModifiedAfterLastAnalyze(NotebookIterator.PageInfo page)
        {
            XAttribute lastModifiedDateAttribute = page.PageElement.Attribute("lastModifiedTime");
            if (lastModifiedDateAttribute != null)
            {
                DateTime lastModifiedDate = DateTime.Parse(lastModifiedDateAttribute.Value);

                string lastAnalyzeTime = OneNoteUtils.GetElementMetaData(page.PageElement, Constants.Key_LatestAnalyzeTime, page.Xnm);
                if (!string.IsNullOrEmpty(lastAnalyzeTime) && lastModifiedDate <= DateTime.Parse(lastAnalyzeTime).ToLocalTime())
                    return false;
            }

            return true;
        }
    }
}
