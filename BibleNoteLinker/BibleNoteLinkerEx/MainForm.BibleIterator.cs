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

namespace BibleNoteLinkerEx
{
    public partial class MainForm
    {
        private const int StagesCount = 5;

        private void StartAnalyze()
        {
            int pagesCount = 0;

            if (!rbAnalyzeCurrentPage.Checked)
            {
                List<NotebookIterator.NotebookInfo> notebooks = GetNotebooksInfo();
                pagesCount = notebooks.Sum(notebook => notebook.PagesCount);

                pbMain.Maximum = pagesCount;
                Logger.LogMessage(Helper.GetRightFoundPagesString(pagesCount));

                foreach (NotebookIterator.NotebookInfo notebook in notebooks)
                    ProcessNotebook(notebook);
            }
            else
            {
                var currentPage = OneNoteUtils.GetCurrentPageInfo(_oneNoteApp);
                if (currentPage != null)
                {
                    string message = "Обработка текущей страницы";

                    pbMain.Maximum = 100;
                    pbMain.Value = 0;

                    LogHighLevelMessage(message, 1, StagesCount);
                    Logger.LogMessage(message);
                    pagesCount = 1;
                    Logger.MoveLevel(1);
                    ProcessPage(currentPage);
                    Logger.MoveLevel(-1);
                }
            }

            if (pagesCount > 0)
            {
                CommitNotesPages();

                UpdateLinksToNotesPages();

                CommitAllPages();

                SortNotesPages();

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
            string message = "Обновление иерархии в OneNote";
            LogHighLevelMessage(message, 5, StagesCount);
            Logger.LogMessage(message, true, false);
            OneNoteProxy.Instance.CommitAllModifiedHierarchy(_oneNoteApp,
                pagesCount =>
                {
                    pbMain.Maximum = pagesCount;
                    pbMain.Value = 0;
                    pbMain.PerformStep();
                    Logger.LogMessage(string.Format(" ({0})", Helper.GetRightPagesString(pagesCount)), false, true, false);
                },
                pageContent => PerformProcessStep());            
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
                    Logger.LogError(string.Format("Ошибка во время сортировки страницы '{0}'", sortPageInfo.PageId), ex);
                }
            }
        }

        private void CommitAllPages()
        {
            string message = "Обновление страниц в OneNote";
            LogHighLevelMessage(message, 4, StagesCount);
            Logger.LogMessage(message, true, false);
            OneNoteProxy.Instance.CommitAllModifiedPages(_oneNoteApp,
                null,
                pagesCount =>
                {
                    pbMain.Maximum = pagesCount;
                    pbMain.Value = 0;
                    pbMain.PerformStep();
                    Logger.LogMessage(string.Format(" ({0})", Helper.GetRightPagesString(pagesCount)), false, true, false);
                },
                pageContent => PerformProcessStep());
        }

        private void UpdateLinksToNotesPages()
        {
            string message = "Обновление ссылок на страницы 'Сводные заметок'";
            LogHighLevelMessage(message, 3, StagesCount);
            Logger.LogMessage(string.Format("{0} ({1})", 
                message, Helper.GetRightPagesString(OneNoteProxy.Instance.ProcessedBiblePages.Values.Count)));            
            pbMain.Maximum = OneNoteProxy.Instance.ProcessedBiblePages.Values.Count;
            pbMain.Value = 0;
            pbMain.PerformStep();
            var relinkNotesManager = new RelinkAllBibleNotesManager(_oneNoteApp);
            foreach (OneNoteProxy.BiblePageId processedBiblePageId in OneNoteProxy.Instance.ProcessedBiblePages.Values)
            {
                relinkNotesManager.RelinkBiblePageNotes(processedBiblePageId.SectionId, processedBiblePageId.PageId,
                    processedBiblePageId.PageName, processedBiblePageId.ChapterPointer);
                PerformProcessStep();
            }
        }

        private void CommitNotesPages()
        {
            string message = "Обновление страниц 'Сводные заметок' в OneNote";
            LogHighLevelMessage(message, 2, StagesCount);
            Logger.LogMessage(message, true, false);
            OneNoteProxy.Instance.CommitAllModifiedPages(_oneNoteApp,
                pageContent => pageContent.PageType == OneNoteProxy.PageType.NotesPage,
                pagesCount =>
                {
                    pbMain.Maximum = pagesCount;
                    pbMain.Value = 0;
                    pbMain.PerformStep();
                    Logger.LogMessage(string.Format(" ({0})", Helper.GetRightPagesString(pagesCount)), false, true, false);
                },
                pageContent => PerformProcessStep());
        }    

        public void ProcessNotebook(NotebookIterator.NotebookInfo notebook)
        {
            if (notebook.PagesCount > 0)
            {
                BibleCommon.Services.Logger.LogMessage("Обработка записной книжки: '{0}'", notebook.Title);
                BibleCommon.Services.Logger.MoveLevel(1);

                ProcessSectionGroup(notebook.RootSectionGroup, true);

                BibleCommon.Services.Logger.MoveLevel(-1);
            }
        }

        private void LogHighLevelMessage(string message, int? stage, int? maxStageCount)
        {
            if (stage.HasValue)
                message = string.Format("Этап {0}/{1}: {2}", stage, maxStageCount, message);

            lblProgress.Text = message;
        }

        private void ProcessSectionGroup(BibleCommon.Services.NotebookIterator.SectionGroupInfo sectionGroup, bool isRoot)
        {
            if (!isRoot)
            {
                BibleCommon.Services.Logger.LogMessage("Обработка группы секций '{0}'", sectionGroup.Title);
                BibleCommon.Services.Logger.MoveLevel(1);
            }

            foreach (BibleCommon.Services.NotebookIterator.SectionInfo section in sectionGroup.Sections)
            {
                BibleCommon.Services.Logger.LogMessage("Обработка секции '{0}'", section.Title);
                BibleCommon.Services.Logger.MoveLevel(1);

                foreach (BibleCommon.Services.NotebookIterator.PageInfo page in section.Pages)
                {
                    string message = string.Format("Обработка страницы '{0}'", page.Title.Replace("{", "{{").Replace("}", "}}"));
                    LogHighLevelMessage(message, 1, StagesCount);
                    BibleCommon.Services.Logger.LogMessage(message);
                    BibleCommon.Services.Logger.MoveLevel(1);

                    ProcessPage(page);                    

                    BibleCommon.Services.Logger.MoveLevel(-1);
                }

                BibleCommon.Services.Logger.MoveLevel(-1);
            }

            foreach (BibleCommon.Services.NotebookIterator.SectionGroupInfo subSectionGroup in sectionGroup.SectionGroups)
            {
                ProcessSectionGroup(subSectionGroup, false);
            }

            if (!isRoot)
                BibleCommon.Services.Logger.MoveLevel(-1);
        }

        private void ProcessPage(NotebookIterator.PageInfo page)
        {
            NoteLinkManager noteLinkManager = new NoteLinkManager(_oneNoteApp);
            noteLinkManager.OnNextVerseProcess += new EventHandler<NoteLinkManager.ProcessVerseEventArgs>(noteLinkManager_OnNextVerseProcess);
            noteLinkManager.LinkPageVerses(page.SectionGroupId, page.SectionId, page.Id, NoteLinkManager.AnalyzeDepth.Full, chkForce.Checked);
            pbMain.PerformStep();
            System.Windows.Forms.Application.DoEvents();

            if (_processAbortedByUser)
                throw new ProcessAbortedByUserException();
        }

        void noteLinkManager_OnNextVerseProcess(object sender, NoteLinkManager.ProcessVerseEventArgs e)
        {
            System.Windows.Forms.Application.DoEvents();
            e.CancelProcess = _processAbortedByUser;

            if (rbAnalyzeCurrentPage.Checked && e.FoundVerse)
            {
                if (pbMain.Value == pbMain.Maximum)
                    pbMain.Value = 0;
                pbMain.PerformStep();
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
