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
        private static Dictionary<string, HashSet<string>> _processedNodes = new Dictionary<string, HashSet<string>>();  // список актуализированных узлов в рамках текущей сессии анализа заметок

        public static bool UpdateNotesPage(ref Application oneNoteApp, NoteLinkManager noteLinkManager,
            VersePointer vp, decimal verseWeight, XmlCursorPosition versePosition, bool isChapter,
            HierarchySearchManager.HierarchyObjectInfo verseHierarchyObjectInfo, 
            HierarchyElementInfo notePageInfo, string notePageContentObjectId, NoteLinkManager.NotesPageType notesPageType, string notesPageName,
            bool isImportantVerse, bool force, bool processAsExtendedVerse)
        {
            if (verseHierarchyObjectInfo.VerseNumber.HasValue)
                vp.VerseNumber = verseHierarchyObjectInfo.VerseNumber.Value;

            var notesPageFilePath = OpenNotesPageHandler.GetNotesPageFilePath(vp, notesPageType);             
            var notesPageData = OneNoteProxy.Instance.GetNotesPageData(notesPageFilePath, notesPageName, vp.IsChapter ? vp : vp.GetChapterPointer());

            var verseNotesPageData = notesPageData.GetVerseNotesPageData(vp);

            AddNotePageLink(ref oneNoteApp, notesPageFilePath, notesPageName, verseNotesPageData, notePageInfo, notePageContentObjectId, verseWeight, versePosition, vp, noteLinkManager, force, processAsExtendedVerse);

            return notesPageData.IsNew;
        }

        private static void AddNotePageLink(ref Application oneNoteApp, string notesPageFilePath, string notesPageName, VerseNotesPageData verseNotesPageData,
            HierarchyElementInfo notePageInfo, string notePageContentObjectId, decimal verseWeight, XmlCursorPosition versePosition, 
            VersePointer vp, NoteLinkManager noteLinkManager, bool force, bool processAsExtendedVerse)
        {            
            NotesPageHierarchyLevelBase parentLevel = verseNotesPageData;
            if (notePageInfo.Parent != null)
                parentLevel = CreateParentTreeStructure(ref oneNoteApp, verseNotesPageData, notePageInfo.Parent, notePageInfo.NotebookId, notesPageFilePath, notesPageName, vp);

            var pageLinkLevel = SearchPageLinkLevel(ref oneNoteApp, notePageInfo, (NotesPageHierarchyLevel)parentLevel, notesPageName, vp, noteLinkManager, force, processAsExtendedVerse);              // parentLevel точно будет типа NotesPageHierarchyLevel

            if (pageLinkLevel == null)
            {
                pageLinkLevel = new NotesPagePageLevel() 
                { 
                    Id = notePageInfo.UniqueName, 
                    Title = notePageInfo.UniqueTitle, 
                    PageId = notePageInfo.Id, 
                    PageTitleObjectId = notePageInfo.UniqueNoteTitleId
                };
                parentLevel.AddLevel(ref oneNoteApp, pageLinkLevel, notePageInfo, true);                
            }
            else if (string.IsNullOrEmpty(pageLinkLevel.GetPageTitleLinkHref(ref oneNoteApp)))  // а то, если десериализовали объект, в котором был только один данный стих, то нет ссылки на заголовок заметки
            {
                pageLinkLevel.SetPageTitleLinkHref(notePageInfo.Id, notePageInfo.UniqueNoteTitleId);
            }

            if (!processAsExtendedVerse)
            {
                pageLinkLevel.AddPageLink(new NotesPageLink()
                                            {
                                                VersePosition = versePosition,
                                                VerseWeight = verseWeight,
                                                PageId = notePageInfo.Id,
                                                ContentObjectId = notePageContentObjectId
                                            }, vp);
            }
        }

        private static NotesPagePageLevel SearchPageLinkLevel(ref Application oneNoteApp, HierarchyElementInfo notePageInfo, NotesPageHierarchyLevel parentLevel, string notesPageName,
            VersePointer vp, NoteLinkManager noteLinkManager, bool force, bool processAsExtendedVerse)
        {
            NotesPagePageLevel pageLinkLevel = null;

            var id = notePageInfo.UniqueName;
            if (parentLevel.Levels.ContainsKey(id))
                pageLinkLevel = (NotesPagePageLevel)parentLevel.Levels[id];

            if (pageLinkLevel == null && parentLevel.Root.AllPagesLevels.ContainsKey(id))
            {
                pageLinkLevel = parentLevel.Root.AllPagesLevels[id];
                pageLinkLevel.Parent.Levels.Remove(pageLinkLevel.Id);
                parentLevel.AddLevel(ref oneNoteApp, pageLinkLevel, notePageInfo, true);
            }

            if (pageLinkLevel != null)
            {
                var key = new NoteLinkManager.NotePageProcessedVerseId() { NotePageId = id, NotesPageName = notesPageName };
                if (force && !noteLinkManager.ContainsNotePageProcessedVerse(key, vp) && !processAsExtendedVerse)  // если в первый раз и force и не расширенный стих
                {  // удаляем старые ссылки на текущую страницу, так как мы начали новый анализ с параметром "force" и мы только в первый раз зашли сюда
                    pageLinkLevel.Parent.Levels.Remove(pageLinkLevel.Id);
                    pageLinkLevel = null;
                }
            }

            if (pageLinkLevel != null)            
                TryToSortElementInParent(ref oneNoteApp, pageLinkLevel, notePageInfo, notesPageName, vp);            

            return pageLinkLevel;
        }

        private static NotesPageHierarchyLevelBase CreateParentTreeStructure(ref Application oneNoteApp, NotesPageHierarchyLevelBase parentLevel, HierarchyElementInfo hierarchyElementInfo,
            string notebookId, string notesPageFilePath, string notesPageName, VersePointer vp)
        {
            NotesPageHierarchyLevelBase parent = parentLevel;
            if (hierarchyElementInfo.Parent != null)
                parentLevel = CreateParentTreeStructure(ref oneNoteApp, parentLevel, hierarchyElementInfo.Parent, notebookId, notesPageFilePath, notesPageName, vp);

            if (!parentLevel.Levels.ContainsKey(hierarchyElementInfo.UniqueName))
            {
                var notesPageLevel = new NotesPageHierarchyLevel() 
                { 
                    Id = hierarchyElementInfo.UniqueName, 
                    Title = hierarchyElementInfo.Title, 
                    HierarchyType = hierarchyElementInfo.Type,
                    OneNoteId = hierarchyElementInfo.Id
                };
                parentLevel.AddLevel(ref oneNoteApp, notesPageLevel, hierarchyElementInfo, true);
                return notesPageLevel;
            }
            else
            {
                var notesPageLevel = parentLevel.Levels[hierarchyElementInfo.UniqueName];

                TryToSortElementInParent(ref oneNoteApp, notesPageLevel, hierarchyElementInfo, notesPageName, vp);                

                return notesPageLevel;
            }
        }

        private static void TryToSortElementInParent(ref Application oneNoteApp, 
            NotesPageHierarchyLevel levelToSort, HierarchyElementInfo hierarchyElementInfo, string notesPageName, VersePointer vp)
        {
            string processedNodeKey = GetProccessNodeKey(notesPageName, vp);
            if (!_processedNodes.ContainsKey(processedNodeKey))
                _processedNodes.Add(processedNodeKey, new HashSet<string>());

            if (!_processedNodes[processedNodeKey].Contains(hierarchyElementInfo.Id))
            {
                levelToSort.TryToSortInParent(ref oneNoteApp, hierarchyElementInfo);

                _processedNodes[processedNodeKey].Add(hierarchyElementInfo.Id);
            }
        }

        private static string GetProccessNodeKey(string notesPageName, VersePointer vp)
        {
            return string.Format("{0}_{1}", notesPageName, vp);
        }
    }
}
