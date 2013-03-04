using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BibleCommon.Common;
using System.IO;
using Microsoft.Office.Interop.OneNote;

namespace BibleCommon.Services
{
    public static class NotesPageManagerFS
    {
        private static Dictionary<string, HashSet<string>> _processedNodes = new Dictionary<string, HashSet<string>>();  // список актуализированных узлов в рамках текущей сессии анализа заметок

        public static void UpdateNotesPage(ref Microsoft.Office.Interop.OneNote.Application oneNoteApp, NoteLinkManager noteLinkManager,
            Common.VersePointer vp, decimal verseWeight, XmlCursorPosition versePosition,
            bool isChapter, HierarchySearchManager.HierarchyObjectInfo verseHierarchyObjectInfo, Common.HierarchyElementInfo notePageInfo, string notePageContentObjectId,
            bool isImportantVerse, bool force, bool processAsExtendedVerse)
        {            
            var notesPageFilePath = GetNotesPageFilePath(vp);
            var notesPageData = OneNoteProxy.Instance.GetNotesPageData(notesPageFilePath);

            var verseNotesPageData = notesPageData.GetVerseNotesPageData(vp);

            AddNotePageLink(ref oneNoteApp, verseNotesPageData, notePageInfo, notesPageFilePath);
        }

        private static void AddNotePageLink(ref Application oneNoteApp, VerseNotesPageData verseNotesPageData, HierarchyElementInfo notePageInfo,
            string notesPageFilePath)
        {
            var 
            if (notePageInfo.Parent != null)
                var parent = CreateParentTreeStructure(ref oneNoteApp, verseNotesPageData, notePageInfo.Parent, notePageInfo.NotebookId, notesPageFilePath);            
        }

        private static NotesPageLevel CreateParentTreeStructure(ref Application oneNoteApp, VerseNotesPageData verseNotesPageData, HierarchyElementInfo hierarchyElementInfo, 
            string notebookId, string notesPageFilePath)
        {
            if (hierarchyElementInfo.Parent != null)
                CreateParentTreeStructure(ref oneNoteApp, verseNotesPageData, hierarchyElementInfo.Parent, notebookId, notesPageFilePath);

            if (!verseNotesPageData.Levels.ContainsKey(hierarchyElementInfo.UniqueName))
            {
                var notesPageLevel = new NotesPageLevel() { ID = hierarchyElementInfo.UniqueName, Title = hierarchyElementInfo.Title };
                verseNotesPageData.Levels.Add(hierarchyElementInfo.UniqueName, notesPageLevel);
            }
            else
            {
            }
        }

        private static string GetNotesPageFilePath(VersePointer vp)
        {
            var path =
                    Path.Combine(
                            Path.Combine(SettingsManager.Instance.FolderPath_BibleNotesPages, SettingsManager.Instance.ModuleShortName),
                            Path.Combine(vp.Book.Name, vp.Chapter.Value.ToString())
                            );

            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);

            return path;
        }
    }
}
