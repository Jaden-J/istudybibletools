using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BibleCommon.Contracts;
using BibleCommon.Helpers;

namespace BibleCommon.Services
{
    public class NotesPagesProviderManager : INotesPageManager
    {
        public string ManagerName
        {
            get { return null; }
        }

        private bool _warningWasLogged = false;

        public Dictionary<string, INotesPageManager> _managers = new Dictionary<string, INotesPageManager>();
        List<INotesPageManager> _registeredManagers = new List<INotesPageManager>() { new NotesPageManager(), new NotesPageManagerEx() };
        public INotesPageManager _defaultManager;

        public NotesPagesProviderManager()
        {
            _defaultManager = new NotesPageManager();
            _managers = new Dictionary<string, INotesPageManager>();

            foreach (var manager in _registeredManagers)
                _managers.Add(manager.ManagerName, manager);
        }

        public string UpdateNotesPage(ref Microsoft.Office.Interop.OneNote.Application oneNoteApp, NoteLinkManager noteLinkManager, Common.VersePointer vp, bool isChapter, 
            HierarchySearchManager.HierarchyObjectInfo verseHierarchyObjectInfo, Common.HierarchyElementInfo notePageId, string notesPageId, string notePageContentObjectId, 
            string notesPageName, int notesPageWidth, bool force, bool processAsExtendedVerse, bool commonNotesPage, out bool rowWasAdded)
        {
            var manager = GetNotesPageProvider(ref oneNoteApp, notesPageId);

            //вставка!! удалить
            if (commonNotesPage)
                manager = _defaultManager;

            return manager.UpdateNotesPage(ref oneNoteApp, noteLinkManager, vp, isChapter, verseHierarchyObjectInfo, notePageId, notesPageId,
                notePageContentObjectId, notesPageName, notesPageWidth, force, processAsExtendedVerse, commonNotesPage, out rowWasAdded);
        }

        public string GetNotesRowObjectId(ref Microsoft.Office.Interop.OneNote.Application oneNoteApp, string notesPageId, Common.VerseNumber? verseNumber, bool isChapter)
        {
            var manager = GetNotesPageProvider(ref oneNoteApp, notesPageId);
            return manager.GetNotesRowObjectId(ref oneNoteApp, notesPageId, verseNumber, isChapter);
        }

        private INotesPageManager GetNotesPageProvider(ref Microsoft.Office.Interop.OneNote.Application oneNoteApp, string notesPageId)
        {
            var notesPageDocument = OneNoteProxy.Instance.GetPageContent(ref oneNoteApp, notesPageId, OneNoteProxy.PageType.NotesPage);

            var pageMetadata = OneNoteUtils.GetElementMetaData(notesPageDocument.Content.Root, Consts.Constants.Key_NotesPageManagerName, notesPageDocument.Xnm);

            if (pageMetadata == null)
            {
                if (!_warningWasLogged)
                {
                    Logger.LogWarning(Resources.Constants.NotesPagesManagerIsObsolete);
                    _warningWasLogged = true;
                }

                return _defaultManager;
            }
            else
                return _managers[pageMetadata];
        }        
    }
}
