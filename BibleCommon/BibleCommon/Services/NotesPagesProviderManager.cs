using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BibleCommon.Contracts;
using BibleCommon.Helpers;
using System.Xml.XPath;
using System.Xml.Linq;
using BibleCommon.Consts;
using BibleCommon.Common;

namespace BibleCommon.Services
{
    public class NotesPagesProviderManager : INotesPageManager
    {
        private bool _warningWasLogged = false;
        private XNamespace _nms;

        public bool ForceUpdateProvider { get; set; }  // если force и анализируем все страницы, то обновляем провайдер и очищаем старые страницы
        public Dictionary<string, INotesPageManager> _managers = new Dictionary<string, INotesPageManager>();
        List<INotesPageManager> _registeredManagers = new List<INotesPageManager>() { new NotesPageManager(), new NotesPageManagerEx() };
        public INotesPageManager _defaultManager;

        public string ManagerName
        {
            get { return null; }
        }

        public NotesPagesProviderManager()
        {
            _nms = XNamespace.Get(Constants.OneNoteXmlNs);
            _defaultManager = new NotesPageManager();
            _managers = new Dictionary<string, INotesPageManager>();

            foreach (var manager in _registeredManagers)
                _managers.Add(manager.ManagerName, manager);
        }

        public string UpdateNotesPage(ref Microsoft.Office.Interop.OneNote.Application oneNoteApp, NoteLinkManager noteLinkManager, 
            Common.VersePointer vp, decimal verseWeight, XmlCursorPosition versePosition,
            bool isChapter, BibleHierarchyObjectInfo verseHierarchyObjectInfo, Common.HierarchyElementInfo notePageId, string notesPageId, string notePageContentObjectId, 
            string notesPageName, int notesPageWidth, bool isImportantVerse, bool force, bool processAsExtendedVerse, out bool rowWasAdded)
        {
            var manager = GetNotesPageProvider(ref oneNoteApp, notesPageId, notesPageWidth);            

            return manager.UpdateNotesPage(ref oneNoteApp, noteLinkManager, vp, verseWeight, versePosition, isChapter, verseHierarchyObjectInfo, notePageId, notesPageId,
                notePageContentObjectId, notesPageName, notesPageWidth, isImportantVerse, force, processAsExtendedVerse, out rowWasAdded);
        }

        public string GetNotesRowObjectId(ref Microsoft.Office.Interop.OneNote.Application oneNoteApp, string notesPageId, Common.VerseNumber? verseNumber, bool isChapter)
        {
            var manager = GetNotesPageProvider(ref oneNoteApp, notesPageId, null);
            return manager.GetNotesRowObjectId(ref oneNoteApp, notesPageId, verseNumber, isChapter);
        }

        private INotesPageManager GetNotesPageProvider(ref Microsoft.Office.Interop.OneNote.Application oneNoteApp, string notesPageId, int? notesPageWidth)
        {
            var notesPageDocument = OneNoteProxy.Instance.GetPageContent(ref oneNoteApp, notesPageId, OneNoteProxy.PageType.NotesPage);

            var pageMetadata = OneNoteUtils.GetElementMetaData(notesPageDocument.Content.Root, Consts.Constants.Key_NotesPageManagerName, notesPageDocument.Xnm);

            if (pageMetadata == null)
            {
                if (ForceUpdateProvider)
                {
                    OneNoteUtils.UpdateElementMetaData(notesPageDocument.Content.Root, Consts.Constants.Key_NotesPageManagerName, NotesPageManagerEx.Const_ManagerName, notesPageDocument.Xnm);
                    foreach (var outlineEl in notesPageDocument.Content.Root.XPathSelectElements("one:Outline", notesPageDocument.Xnm))
                    {
                        outlineEl.RemoveNodes();

                        if (notesPageWidth.HasValue)
                            outlineEl.Add(new XElement(_nms + "Size", new XAttribute("width", notesPageWidth), new XAttribute("height", 15), new XAttribute("isSetByUser", true)));
                    }

                    return _managers[NotesPageManagerEx.Const_ManagerName];
                }
                else if (!_warningWasLogged)
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
