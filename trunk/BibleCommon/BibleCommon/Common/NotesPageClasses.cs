using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using BibleCommon.Services;

namespace BibleCommon.Common
{
    public class NotesPageData
    {
        public string FilePath { get; set; }
        public Dictionary<VersePointer, VerseNotesPageData> VersesNotesPageData { get; set; }                

        public NotesPageData(string filePath)
        {
            this.FilePath = filePath;

            if (File.Exists(this.FilePath))
                Deserialize();
        }

        public void Update(ref Microsoft.Office.Interop.OneNote.Application oneNoteApp, NoteLinkManager noteLinkManager,
            Common.VersePointer vp, decimal verseWeight, XmlCursorPosition versePosition,
            bool isChapter, HierarchySearchManager.HierarchyObjectInfo verseHierarchyObjectInfo, Common.HierarchyElementInfo notePageId, string notesPageId, string notePageContentObjectId,
            string notesPageName, int notesPageWidth, bool isImportantVerse, bool force, bool processAsExtendedVerse, out bool rowWasAdded)
        {
            rowWasAdded = true;
        }

        protected void Deserialize()
        {
            throw new NotImplementedException();
        }

        public void Serialize()
        {

        }
    }

    public class VerseNotesPageData
    {
        public VersePointer Verse { get; set; }
        public List<NotesPageLevel> Levels { get; set; }
        public Dictionary<string, NotesPageLevel> AllPagesLevels { get; set; }
    }
    

    public class NotesPageLevel
    {
        public string Title { get; set; }
        public string Href { get; set; }
        public string ID { get; set; }

        public NotesPageLevel Parent { get; set; }
        public List<NotesPageLevel> ChildLevels { get; set; }

        private bool _pageLinksWasParsed = false;
        private List<NotesPageLink> _pageLinks;
        public List<NotesPageLink> PageLinks 
        {
            get
            {
                if (!_pageLinksWasParsed)
                {
                    PageLinks.ForEach(pl => pl.Parse());
                    _pageLinksWasParsed = true;
                }

                return _pageLinks;
            }
        }
    }

    public class NotesPageLink
    {
        public string Href { get; set; }
        public int Position { get; set; }
        internal bool WasParsed { get; set; }

        internal void Parse()
        {
            //Position = ;
            WasParsed = true;
        }
    }
}
