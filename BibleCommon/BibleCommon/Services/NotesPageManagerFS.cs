using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BibleCommon.Common;
using System.IO;
using Microsoft.Office.Interop.OneNote;
using BibleCommon.Helpers;
using BibleCommon.Handlers;

namespace BibleCommon.Services
{
    public static class NotesPageManagerFS
    {
        //private static Dictionary<string, HashSet<string>> _processedNodes = new Dictionary<string, HashSet<string>>();  // список актуализированных узлов в рамках текущей сессии анализа заметок

        public static bool UpdateNotesPage(ref Application oneNoteApp, NoteLinkManager noteLinkManager,
            VersePointer vp, decimal verseWeight, XmlCursorPosition versePosition, bool isChapter,
            HierarchySearchManager.HierarchyObjectInfo verseHierarchyObjectInfo, 
            HierarchyElementInfo notePageInfo, string notePageContentObjectId, NoteLinkManager.NotesPageType notesPageType, string notesPageName,
            bool isImportantVerse, bool force, bool processAsExtendedVerse)
        {
            if (verseHierarchyObjectInfo.VerseNumber.HasValue)
                vp.VerseNumber = verseHierarchyObjectInfo.VerseNumber.Value;

            var notesPageFilePath = OpenNotesPageHandler.GetNotesPageFilePath(vp, notesPageType);             
            var notesPageData = OneNoteProxy.Instance.GetNotesPageData(notesPageFilePath, notesPageName);

            var verseNotesPageData = notesPageData.GetVerseNotesPageData(vp);

            AddNotePageLink(ref oneNoteApp, notesPageFilePath, verseNotesPageData, notePageInfo, notePageContentObjectId, verseWeight, versePosition, vp, noteLinkManager, force, processAsExtendedVerse);

            return notesPageData.IsNew;
        }

        private static void AddNotePageLink(ref Application oneNoteApp, string notesPageFilePath, VerseNotesPageData verseNotesPageData,
            HierarchyElementInfo notePageInfo, string notePageContentObjectId, decimal verseWeight, XmlCursorPosition versePosition, 
            VersePointer vp, NoteLinkManager noteLinkManager, bool force, bool processAsExtendedVerse)
        {            
            NotesPageHierarchyLevelBase parentLevel = verseNotesPageData;
            if (notePageInfo.Parent != null)
                parentLevel = CreateParentTreeStructure(ref oneNoteApp, verseNotesPageData, notePageInfo.Parent, notePageInfo.NotebookId, notesPageFilePath);

            var pageLinkLevel = SearchPageLinkLevel(notePageInfo.UniqueName, (NotesPageHierarchyLevel)parentLevel, notesPageFilePath, vp, noteLinkManager, force, processAsExtendedVerse);              // parentLevel точно будет типа NotesPageHierarchyLevel

            if (pageLinkLevel == null)
            {
                pageLinkLevel = new NotesPagePageLevel() { ID = notePageInfo.UniqueName, Title = notePageInfo.UniqueTitle, PageId = notePageInfo.Id, PageTitleObjectId = notePageInfo.UniqueNoteTitleId };
                parentLevel.AddLevel(pageLinkLevel);                
            }

            pageLinkLevel.AddPageLink(
                        new NotesPageLink()
                        {
                            VersePosition = versePosition,
                            VerseWeight = verseWeight,
                            PageId = notePageInfo.Id,
                            ContentObjectId = notePageContentObjectId                            
                        },
                        vp);
        }

        private static NotesPagePageLevel SearchPageLinkLevel(string id, NotesPageHierarchyLevel parentLevel, string notesPageFilePath,
            VersePointer vp, NoteLinkManager noteLinkManager, bool force, bool processAsExtendedVerse)
        {
            NotesPagePageLevel pageLinkLevel = null;

            if (parentLevel.Levels.ContainsKey(id))
                pageLinkLevel = (NotesPagePageLevel)parentLevel.Levels[id];

            if (pageLinkLevel == null && parentLevel.Root.AllPagesLevels.ContainsKey(id))
            {
                pageLinkLevel = parentLevel.Root.AllPagesLevels[id];
                pageLinkLevel.Parent.Levels.Remove(pageLinkLevel.ID);
                parentLevel.AddLevel(pageLinkLevel);                
            }

            if (pageLinkLevel != null)
            {
                var key = new NoteLinkManager.NotePageProcessedVerseId() { NotePageId = id, NotesPageName = notesPageFilePath };
                if (force && !noteLinkManager.ContainsNotePageProcessedVerse(key, vp) && !processAsExtendedVerse)  // если в первый раз и force и не расширенный стих
                {  // удаляем старые ссылки на текущую страницу, так как мы начали новый анализ с параметром "force" и мы только в первый раз зашли сюда
                    pageLinkLevel.Parent.Levels.Remove(pageLinkLevel.ID);
                    pageLinkLevel = null;
                }
            }

            if (pageLinkLevel != null)
            {
                //todo: sort once
            }

            return pageLinkLevel;
        }

        private static NotesPageHierarchyLevelBase CreateParentTreeStructure(ref Application oneNoteApp, NotesPageHierarchyLevelBase parentLevel, HierarchyElementInfo hierarchyElementInfo, 
            string notebookId, string notesPageFilePath)
        {
            NotesPageHierarchyLevelBase parent = parentLevel;
            if (hierarchyElementInfo.Parent != null)
                parentLevel = CreateParentTreeStructure(ref oneNoteApp, parentLevel, hierarchyElementInfo.Parent, notebookId, notesPageFilePath);

            if (!parentLevel.Levels.ContainsKey(hierarchyElementInfo.UniqueName))
            {
                var notesPageLevel = new NotesPageHierarchyLevel() { ID = hierarchyElementInfo.UniqueName, Title = hierarchyElementInfo.Title };
                parentLevel.AddLevel(notesPageLevel);
                return notesPageLevel;
            }
            else
            {
                //todo: sort once
                return parentLevel.Levels[hierarchyElementInfo.UniqueName];
            }
        }       
    }
}
