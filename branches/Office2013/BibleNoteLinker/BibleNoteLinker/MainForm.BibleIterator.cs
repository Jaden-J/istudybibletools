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

namespace BibleNoteLinker
{
    public partial class MainForm
    {
        private const int StagesCount = 5;
        private const int ApproximatePageVersesCount = 100;
        private int _pagesForAnalyzeCount;

        private void StartAnalyze()
        {            
            pbMain.Value = 0;

            try
            {
                OneNoteLocker.UnlockAllBible(_oneNoteApp);
            }
            catch (NotSupportedException)
            {
                //todo: log it
            }

            if (!rbAnalyzeCurrentPage.Checked)
            {
                List<NotebookIterator.NotebookInfo> notebooks = GetNotebooksInfo();
                _pagesForAnalyzeCount = notebooks.Sum(notebook => notebook.PagesCount);

                pbMain.Maximum = _pagesForAnalyzeCount > 1 ? _pagesForAnalyzeCount : ApproximatePageVersesCount;

                pbMain.PerformStep();
                Logger.LogMessage(Helper.GetRightFoundPagesString(_pagesForAnalyzeCount));

                foreach (NotebookIterator.NotebookInfo notebook in notebooks)
                    ProcessNotebook(notebook);
            }
            else
            {
                var currentPage = OneNoteUtils.GetCurrentPageInfo(_oneNoteApp);

                _pagesForAnalyzeCount = 1;
                string message = BibleCommon.Resources.Constants.ProcessCurrentPage;

                pbMain.Maximum = ApproximatePageVersesCount;

                LogHighLevelMessage(message, 1, StagesCount);
                Logger.LogMessage(message);
                Logger.MoveLevel(1);
                ProcessPage(currentPage);
                Logger.MoveLevel(-1);
            }

            if (_pagesForAnalyzeCount > 0)
            {
                CommitNotesPages();

                UpdateLinksToNotesPages();

                CommitAllPages();

                SortNotesPages();  // это происходит очень быстро, поэтому не выделяем это как отдельный этап

                CommitNotesPagesHierarchy();
            }          
        }

        private void PerformProcessStep()
        {
            System.Windows.Forms.Application.DoEvents();
            if (_processAbortedByUser)
                throw new ProcessAbortedByUserException();
            pbMain.PerformStep();
        }

        private void CommitNotesPagesHierarchy()
        {
            string message = BibleCommon.Resources.Constants.NoteLilnkerHierarchyUpdating;
            LogHighLevelMessage(message, 5, StagesCount);
            int allPagesCount = 0;
            int processedPagesCount = 0;
            Logger.LogMessage(message, true, false);
            OneNoteProxy.Instance.CommitAllModifiedHierarchy(_oneNoteApp,
                pagesCount =>
                {
                    allPagesCount = pagesCount;
                    pbMain.Maximum = pagesCount;
                    pbMain.Value = 0;
                    pbMain.PerformStep();
                    Logger.LogMessage(string.Format(" ({0})", Helper.GetRightPagesString(pagesCount)), false, true, false);
                    LogHighLevelAdditionalMessage(string.Format(": {0}/{1}", ++processedPagesCount, allPagesCount));
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
            foreach (var sortPageInfo in OneNoteProxy.Instance.SortVerseLinkPagesInfo)
            {
                try
                {
                    VerseLinkManager.SortVerseLinkPages(_oneNoteApp,
                        sortPageInfo.SectionId, sortPageInfo.PageId, sortPageInfo.ParentPageId, sortPageInfo.PageLevel);
                }
                catch (Exception ex)
                {
                    Logger.LogError(string.Format("{0} '{1}'", BibleCommon.Resources.Constants.NoteLinkerErrorWhilePageSorting, sortPageInfo.PageId), ex);
                }
            }
        }

        private void CommitAllPages()
        {
            string message = BibleCommon.Resources.Constants.NoteLinkerPagesUpdating;
            LogHighLevelMessage(message, 4, StagesCount);
            Logger.LogMessage(message, true, false);
            int allPagesCount = 0;
            int processedPagesCount = 0;
            OneNoteProxy.Instance.CommitAllModifiedPages(_oneNoteApp,
                null,
                pagesCount =>
                {
                    allPagesCount = pagesCount;
                    pbMain.Maximum = pagesCount;
                    pbMain.Value = 0;
                    pbMain.PerformStep();
                    Logger.LogMessage(string.Format(" ({0})", Helper.GetRightPagesString(pagesCount)), false, true, false);
                    LogHighLevelAdditionalMessage(string.Format(": {0}/{1}", ++processedPagesCount, allPagesCount));
                },
                pageContent => 
                {
                    PerformProcessStep();
                    LogHighLevelAdditionalMessage(string.Format(": {0}/{1}", ++processedPagesCount, allPagesCount));
                });
        }

        private void UpdateLinksToNotesPages()
        {
            string message = BibleCommon.Resources.Constants.NoteLinkerLinksToNotesPagesUpdating;
            LogHighLevelMessage(message, 3, StagesCount);
            int allPagesCount = OneNoteProxy.Instance.ProcessedBiblePages.Values.Count;            
            Logger.LogMessage(string.Format("{0} ({1})",
                message, Helper.GetRightPagesString(allPagesCount)));            
            pbMain.Maximum = allPagesCount;
            pbMain.Value = 0;
            pbMain.PerformStep();

            int processedPagesCount = 0;
            using (var relinkNotesManager = new RelinkAllBibleNotesManager(_oneNoteApp))
            {
                foreach (OneNoteProxy.BiblePageId processedBiblePageId in OneNoteProxy.Instance.ProcessedBiblePages.Values)
                {
                    LogHighLevelAdditionalMessage(string.Format(": {0}/{1}", ++processedPagesCount, allPagesCount));

                    relinkNotesManager.RelinkBiblePageNotes(processedBiblePageId.SectionId, processedBiblePageId.PageId,
                        processedBiblePageId.PageName, processedBiblePageId.ChapterPointer);

                    PerformProcessStep();                                        
                }
            }
        }

        private void CommitNotesPages()
        {
            string message = BibleCommon.Resources.Constants.NoteLinkerNotesPagesUpdating;
            LogHighLevelMessage(message, 2, StagesCount);
            int allPagesCount = 0;
            int processedPagesCount = 0;
            Logger.LogMessage(message, true, false);
            OneNoteProxy.Instance.CommitAllModifiedPages(_oneNoteApp,
                pageContent => pageContent.PageType == OneNoteProxy.PageType.NotesPage,
                pagesCount =>
                {
                    allPagesCount = pagesCount;
                    pbMain.Maximum = pagesCount;
                    pbMain.Value = 0;
                    pbMain.PerformStep();
                    Logger.LogMessage(string.Format(" ({0})", Helper.GetRightPagesString(pagesCount)), false, true, false);
                    LogHighLevelAdditionalMessage(string.Format(": {0}/{1}", ++processedPagesCount, allPagesCount));
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
                Logger.LogMessage("{0}: '{1}'", BibleCommon.Resources.Constants.NoteLinkerProcessNotebook, notebook.Title);
                Logger.MoveLevel(1);

                ProcessSectionGroup(notebook.RootSectionGroup, true);

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
                message = string.Format("{0} {1}/{2}: {3}", BibleCommon.Resources.Constants.NoteLinkerStage, stage, maxStageCount, message);

            _highLevelMessage = message;
            lblProgress.Text = message;
        }

        private void ProcessSectionGroup(BibleCommon.Services.NotebookIterator.SectionGroupInfo sectionGroup, bool isRoot)
        {
            if (!isRoot)
            {
                Logger.LogMessage("{0} '{1}'", BibleCommon.Resources.Constants.ProcessSectionGroup, sectionGroup.Title);
                Logger.MoveLevel(1);
            }

            foreach (BibleCommon.Services.NotebookIterator.SectionInfo section in sectionGroup.Sections)
            {
                Logger.LogMessage("{0} '{1}'", BibleCommon.Resources.Constants.ProcessSection, section.Title);
                Logger.MoveLevel(1);

                foreach (BibleCommon.Services.NotebookIterator.PageInfo page in section.Pages)
                {
                    string message = string.Format("{0} '{1}'", BibleCommon.Resources.Constants.ProcessPage, page.Title);
                    LogHighLevelMessage(message, 1, StagesCount);
                    Logger.LogMessage(message);
                    Logger.MoveLevel(1);

                    ProcessPage(page);                    

                    Logger.MoveLevel(-1);
                }

                Logger.MoveLevel(-1);
            }

            foreach (BibleCommon.Services.NotebookIterator.SectionGroupInfo subSectionGroup in sectionGroup.SectionGroups)
            {
                ProcessSectionGroup(subSectionGroup, false);
            }

            if (!isRoot)
                Logger.MoveLevel(-1);
        }

        private void ProcessPage(NotebookIterator.PageInfo page)
        {
            using (NoteLinkManager noteLinkManager = new NoteLinkManager(_oneNoteApp))
            {
                noteLinkManager.OnNextVerseProcess += new EventHandler<NoteLinkManager.ProcessVerseEventArgs>(noteLinkManager_OnNextVerseProcess);
                noteLinkManager.LinkPageVerses(page.SectionGroupId, page.SectionId, page.Id, NoteLinkManager.AnalyzeDepth.Full, chkForce.Checked);
            }

            pbMain.PerformStep();
            System.Windows.Forms.Application.DoEvents();            

            if (_processAbortedByUser)
                throw new ProcessAbortedByUserException();
        }

        void noteLinkManager_OnNextVerseProcess(object sender, NoteLinkManager.ProcessVerseEventArgs e)
        {
            System.Windows.Forms.Application.DoEvents();
            e.CancelProcess = _processAbortedByUser;
            
            if (_pagesForAnalyzeCount == 1 && e.FoundVerse)
            {
                if (pbMain.Value == pbMain.Maximum)
                    pbMain.Value = 0;
                pbMain.PerformStep();
            }

            if (e.FoundVerse)
            {
                LogHighLevelAdditionalMessage(string.Format(": {0}", e.VersePointer.OriginalVerseName));
            }
        }

        private List<NotebookIterator.NotebookInfo> GetNotebooksInfo()
        {
            NotebookIterator iterator = new NotebookIterator(_oneNoteApp);
            List<NotebookIterator.NotebookInfo> result = new List<NotebookIterator.NotebookInfo>();

            Func<NotebookIterator.PageInfo, bool> filter = null;
            if (rbAnalyzeChangedPages.Checked)
                filter = IsPageWasModifiedAfterLastAnalyze;

            foreach (string id in Helper.GetSelectedNotebooksIds())
            {
                if (SettingsManager.Instance.IsSingleNotebook)
                    result.Add(iterator.GetNotebookPages(SettingsManager.Instance.NotebookId_Bible, id, filter));
                else
                    result.Add(iterator.GetNotebookPages(id, null, filter));
            }

            return result;
        }

        private bool IsPageWasModifiedAfterLastAnalyze(NotebookIterator.PageInfo page)
        {
            XAttribute lastModifiedDateAttribute = page.PageElement.Attribute("lastModifiedTime");
            if (lastModifiedDateAttribute != null)
            {
                DateTime lastModifiedDate = DateTime.Parse(lastModifiedDateAttribute.Value);

                string lastAnalyzeTime = OneNoteUtils.GetPageMetaData(_oneNoteApp, page.PageElement, Constants.Key_LatestAnalyzeTime, page.Xnm);
                if (!string.IsNullOrEmpty(lastAnalyzeTime) && lastModifiedDate <= DateTime.Parse(lastAnalyzeTime).ToLocalTime())
                    return false;
            }

            return true;
        }
    }
}
