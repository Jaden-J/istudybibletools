using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BibleCommon.Common;
using System.IO;
using Microsoft.Office.Interop.OneNote;
using BibleCommon.Helpers;
using BibleCommon.Handlers;
using System.Reflection;

namespace BibleCommon.Services
{
    public static class NotesPageManagerFS
    {
        private static Dictionary<string, HashSet<string>> _processedNodes = new Dictionary<string, HashSet<string>>();  // список актуализированных узлов в рамках текущей сессии анализа заметок
        private static Dictionary<string, int?> _notebooksDisplayLevel;

        public static bool UpdateNotesPage(ref Application oneNoteApp, NoteLinkManager noteLinkManager,
            VersePointer vp, decimal verseWeight, XmlCursorPosition versePosition, bool isChapter,
            BibleHierarchyObjectInfo verseHierarchyObjectInfo, 
            HierarchyElementInfo notePageInfo, string notePageContentObjectId, NotesPageType notesPageType, string notesPageName,
            bool isImportantVerse, bool force, bool processAsExtendedVerse, bool toDeserializeIfExists, NoteLinkManager.DetailedLink isDetailedLink)
        {
            if (_notebooksDisplayLevel == null)
                LoadNotebooksDisplayLevel();            

            var notesPageFilePath = OpenNotesPageHandler.GetNotesPageFilePath(vp, notesPageType);             
            var notesPageData = ApplicationCache.Instance.GetNotesPageData(notesPageFilePath, notesPageName, notesPageType, vp.IsChapter ? vp : vp.GetChapterPointer(), toDeserializeIfExists);

            var verseNotesPageData = notesPageData.GetVerseNotesPageData(vp);

            AddNotePageLink(ref oneNoteApp, notesPageFilePath, notesPageName, verseNotesPageData, notePageInfo, notePageContentObjectId, verseWeight, versePosition, vp, noteLinkManager, force, processAsExtendedVerse, isDetailedLink);

            return notesPageData.IsNew;
        }               

        private static void AddNotePageLink(ref Application oneNoteApp, string notesPageFilePath, string notesPageName, VerseNotesPageData verseNotesPageData,
            HierarchyElementInfo notePageInfo, string notePageContentObjectId, decimal verseWeight, XmlCursorPosition versePosition,
            VersePointer vp, NoteLinkManager noteLinkManager, bool force, bool processAsExtendedVerse, NoteLinkManager.DetailedLink isDetailedLink)
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

            if (isDetailedLink != NoteLinkManager.DetailedLink.ChangeDetailedOnNotDetailed)
            {
                if (!processAsExtendedVerse || !vp.IsFirstVerseInParentVerse())
                {
                    pageLinkLevel.AddPageLink(new NotesPageLink()
                                                {
                                                    VersePosition = versePosition,
                                                    VerseWeight = verseWeight,
                                                    PageId = notePageInfo.Id,
                                                    ContentObjectId = notePageContentObjectId,
                                                    IsDetailed = isDetailedLink == NoteLinkManager.DetailedLink.Yes,
                                                    VersePointer = vp
                                                }, vp);
                }
            }
            else
            {
                var notePageLink = pageLinkLevel.PageLinks.FirstOrDefault(link => object.ReferenceEquals(link.VersePointer, vp));
                if (notePageLink != null)
                    notePageLink.IsDetailed = false;
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
                if (force && !noteLinkManager.ContainsNotePageProcessedVerse(key, vp) && !processAsExtendedVerse)  // если в первый раз и force и не расширенный стих. Важно: если у нас force, то processAsExtendedVerse будет false
                {
                    pageLinkLevel.Parent.Levels.Remove(pageLinkLevel.Id);                                          // удаляем старые ссылки на текущую страницу, так как мы начали новый анализ с параметром "force" и мы только в первый раз зашли сюда
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
            if (hierarchyElementInfo.Parent != null)
                parentLevel = CreateParentTreeStructure(ref oneNoteApp, parentLevel, hierarchyElementInfo.Parent, notebookId, notesPageFilePath, notesPageName, vp);

            if (!parentLevel.Levels.ContainsKey(hierarchyElementInfo.UniqueName))
            {
                var displayLevel = GetDisplayLevel(notebookId);   

                var notesPageLevel = new NotesPageHierarchyLevel() 
                { 
                    Id = hierarchyElementInfo.UniqueName, 
                    Title = hierarchyElementInfo.Title, 
                    HierarchyType = hierarchyElementInfo.Type,
                    OneNoteId = hierarchyElementInfo.Id,
                    Collapsed = hierarchyElementInfo.GetLevel() == displayLevel.GetValueOrDefault(int.MaxValue)                    
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

        public static void UpdateNotesPageCssFile()
        {
            UpdateNotesPageFile("BibleCommon.Resources.NotesPage.css", Consts.Constants.NotesPageStyleFileName);            
        }

        public static void UpdateNotesPageJsFile()
        {
            UpdateNotesPageFile("BibleCommon.Resources.NotesPage.js", Consts.Constants.NotesPageScriptFileName);
            UpdateNotesPageFile("BibleCommon.Resources.JQuery.js", Consts.Constants.NotesPageJQueryScriptFileName);
        }

        public static void UpdateNotesPageImages()
        {
            var folder = Path.Combine(SettingsManager.Instance.FolderPath_BibleNotesPages, "images");
            if (!Directory.Exists(folder))
                Directory.CreateDirectory(folder);

            UpdateNotesPageFile("BibleCommon.Resources.images.none.png", "images/none.png");
            UpdateNotesPageFile("BibleCommon.Resources.images.minus.png", "images/minus.png");
            UpdateNotesPageFile("BibleCommon.Resources.images.plus.png", "images/plus.png");            
        }

        private static void UpdateNotesPageFile(string fileResourceName, string fileName)
        {
            using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(fileResourceName))
            {
                File.WriteAllBytes(Path.Combine(SettingsManager.Instance.FolderPath_BibleNotesPages, fileName), Utils.ReadStream(stream));
            }
        }

        private static int? GetDisplayLevel(string notebookId)
        {
            int? displayLevel = null;
            if (_notebooksDisplayLevel.ContainsKey(notebookId))
                displayLevel = _notebooksDisplayLevel[notebookId];

            return displayLevel;
        }

        private static void LoadNotebooksDisplayLevel()
        {
            _notebooksDisplayLevel = new Dictionary<string, int?>();
            if (SettingsManager.Instance.SelectedNotebooksForAnalyze != null)
            {
                foreach (var notebookInfo in SettingsManager.Instance.SelectedNotebooksForAnalyze)
                    if (!_notebooksDisplayLevel.ContainsKey(notebookInfo.NotebookId))  // на всякий пожарный
                        _notebooksDisplayLevel.Add(notebookInfo.NotebookId, notebookInfo.DisplayLevels);
            }
        }
    }
}
