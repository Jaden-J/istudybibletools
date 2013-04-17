﻿using System;
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
        public static void Iterate(ref Application oneNoteApp, string notebookId, string sectionGroupId, Action<BibleCommon.Services.NotebookIterator.PageInfo> pageAction)
        {
            if (pageAction == null)
                throw new ArgumentNullException("pageAction");

            var iterator = new NotebookIterator();
            var notebook = iterator.GetNotebookPages(ref oneNoteApp, notebookId, sectionGroupId, null);

            BibleCommon.Services.Logger.LogMessageParams("{0}: '{1}'", BibleCommon.Resources.Constants.ProcessNotebook, notebook.Title);

            BibleCommon.Services.Logger.MoveLevel(1);
            IterateContainer(notebook.RootSectionGroup, true, pageAction);
            BibleCommon.Services.Logger.MoveLevel(-1);

        }

        private static void IterateContainer(BibleCommon.Services.NotebookIterator.SectionGroupInfo sectionGroup, bool isRoot, 
            Action<BibleCommon.Services.NotebookIterator.PageInfo> pageAction)
        {
            if (!isRoot)
            {
                BibleCommon.Services.Logger.LogMessageParams("{0} '{1}'", BibleCommon.Resources.Constants.ProcessSectionGroup, sectionGroup.Title);
                BibleCommon.Services.Logger.MoveLevel(1);
            }

            foreach (BibleCommon.Services.NotebookIterator.SectionInfo section in sectionGroup.Sections)
            {
                BibleCommon.Services.Logger.LogMessageParams("{0} '{1}'", BibleCommon.Resources.Constants.ProcessSection, section.Title);
                BibleCommon.Services.Logger.MoveLevel(1);

                foreach (BibleCommon.Services.NotebookIterator.PageInfo page in section.Pages)
                {
                    BibleCommon.Services.Logger.LogMessageParams("{0} '{1}'", BibleCommon.Resources.Constants.ProcessPage, page.Title);
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
