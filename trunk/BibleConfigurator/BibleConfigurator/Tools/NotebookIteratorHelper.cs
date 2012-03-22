using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BibleCommon.Helpers;
using Microsoft.Office.Interop.OneNote;
using BibleCommon.Services;


namespace BibleConfigurator.Tools
{
    public static class NotebookIteratorHelper
    {
        public static void Iterate(Application oneNoteApp, string notebookId, string sectionGroupId, Action<BibleCommon.Services.NotebookIterator.PageInfo> pageAction)
        {
            if (pageAction == null)
                throw new ArgumentNullException("pageAction");            

            NotebookIterator iterator = new NotebookIterator(oneNoteApp);

            BibleCommon.Services.NotebookIterator.NotebookInfo notebook = iterator.GetNotebookPages(notebookId, sectionGroupId, null);

            BibleCommon.Services.Logger.LogMessage("Обработка записной книжки: '{0}'", notebook.Title);

            BibleCommon.Services.Logger.MoveLevel(1);
            IterateContainer(notebook.RootSectionGroup, true, pageAction);
            BibleCommon.Services.Logger.MoveLevel(-1);
        }

        private static void IterateContainer(BibleCommon.Services.NotebookIterator.SectionGroupInfo sectionGroup, bool isRoot, 
            Action<BibleCommon.Services.NotebookIterator.PageInfo> pageAction)
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
                    BibleCommon.Services.Logger.LogMessage("Обработка страницы '{0}'", page.Title);
                    BibleCommon.Services.Logger.MoveLevel(1);

                    pageAction(page);

                    BibleCommon.Services.Logger.MoveLevel(-1);
                }

                BibleCommon.Services.Logger.MoveLevel(-1);
            }

            foreach (BibleCommon.Services.NotebookIterator.SectionGroupInfo subSectionGroup in sectionGroup.SectionGroups)
            {
                IterateContainer(subSectionGroup, false, pageAction);
            }

            if (!isRoot)
                BibleCommon.Services.Logger.MoveLevel(-1);
        }
    }
}
