using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BibleCommon.Common;
using System.IO;
using Microsoft.Office.Interop.OneNote;
using BibleCommon.Helpers;

namespace BibleCommon.Services
{
    public static class NotesPageManagerFS
    {
        private static Dictionary<string, HashSet<string>> _processedNodes = new Dictionary<string, HashSet<string>>();  // список актуализированных узлов в рамках текущей сессии анализа заметок

        public static bool UpdateNotesPage(ref Application oneNoteApp, NoteLinkManager noteLinkManager,
            VersePointer vp, decimal verseWeight, XmlCursorPosition versePosition,
            bool isChapter, HierarchySearchManager.HierarchyObjectInfo verseHierarchyObjectInfo, HierarchyElementInfo notePageInfo, string notesPageName, string notePageContentObjectId,
            bool isImportantVerse, bool force, bool processAsExtendedVerse)
        {
            var notesPageFilePath = GetNotesPageFilePath(vp, notesPageName); 
            var notesPageData = OneNoteProxy.Instance.GetNotesPageData(notesPageFilePath);

            var verseNotesPageData = notesPageData.GetVerseNotesPageData(vp);

            AddNotePageLink(ref oneNoteApp, notesPageFilePath, verseNotesPageData, notePageInfo, notePageContentObjectId, verseWeight, versePosition, vp, noteLinkManager, force, processAsExtendedVerse);

            return notesPageData.IsNew;
        }

        private static void AddNotePageLink(ref Application oneNoteApp, string notesPageFilePath, VerseNotesPageData verseNotesPageData,
            HierarchyElementInfo notePageInfo, string notePageContentObjectId, decimal verseWeight, XmlCursorPosition versePosition, 
            VersePointer vp, NoteLinkManager noteLinkManager, bool force, bool processAsExtendedVerse)
        {            
            NotesPageLevelBase parentLevel = verseNotesPageData;
            if (notePageInfo.Parent != null)
                parentLevel = CreateParentTreeStructure(ref oneNoteApp, verseNotesPageData, notePageInfo.Parent, notePageInfo.NotebookId, notesPageFilePath);

            var pageLinkLevel = SearchPageLinkLevel(notePageInfo.UniqueName, (NotesPageLevel)parentLevel, notesPageFilePath, vp, noteLinkManager, force, processAsExtendedVerse);              // parentLevel точно будет типа NotesPageLevel

            if (pageLinkLevel == null)
            {
                pageLinkLevel = new NotesPageLevel() { ID = notePageInfo.UniqueName, Title = notePageInfo.UniqueTitle, Type = NotesPageLevelType.Page };
                parentLevel.AddLevel(pageLinkLevel);                
            }

            pageLinkLevel.PageLinks.Add(
                        new NotesPageLink()
                        {
                            VersePosition = versePosition,
                            VerseWeight = verseWeight,
                            Href = OneNoteUtils.GetOrGenerateLinkHref(ref oneNoteApp, null, notePageInfo.Id, notePageContentObjectId, true)
                        });
        }

        private static NotesPageLevel SearchPageLinkLevel(string id, NotesPageLevel parentLevel, string notesPageFilePath,
            VersePointer vp, NoteLinkManager noteLinkManager, bool force, bool processAsExtendedVerse)
        {
            NotesPageLevel pageLinkLevel = null;

            if (parentLevel.Levels.ContainsKey(id))
                pageLinkLevel = parentLevel.Levels[id];

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

        private static NotesPageLevelBase CreateParentTreeStructure(ref Application oneNoteApp, NotesPageLevelBase parentLevel, HierarchyElementInfo hierarchyElementInfo, 
            string notebookId, string notesPageFilePath)
        {
            NotesPageLevelBase parent = parentLevel;
            if (hierarchyElementInfo.Parent != null)
                parentLevel = CreateParentTreeStructure(ref oneNoteApp, parentLevel, hierarchyElementInfo.Parent, notebookId, notesPageFilePath);

            if (!parentLevel.Levels.ContainsKey(hierarchyElementInfo.UniqueName))
            {
                var notesPageLevel = new NotesPageLevel() { ID = hierarchyElementInfo.UniqueName, Title = hierarchyElementInfo.Title, Type = NotesPageLevelType.HierarchyElement };
                parentLevel.AddLevel(notesPageLevel);
                return notesPageLevel;
            }
            else
            {
                //todo: sort once
                return parentLevel.Levels[hierarchyElementInfo.UniqueName];
            }
        }

        private static string GetNotesPageFilePath(VersePointer vp, string notesPageName)
        {
            var path =
                    Path.Combine(
                            Path.Combine(SettingsManager.Instance.FolderPath_BibleNotesPages, SettingsManager.Instance.ModuleShortName),
                            Path.Combine(vp.Book.Name, vp.Chapter.Value.ToString())
                            );

            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);

            return Path.Combine(path, notesPageName + ".htm");
        }
    }
}
