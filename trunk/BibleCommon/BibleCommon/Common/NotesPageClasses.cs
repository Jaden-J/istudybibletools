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
        public bool IsNew { get; set; }
        public string FilePath { get; set; }
        public Dictionary<VersePointer, VerseNotesPageData> VersesNotesPageData { get; set; }                

        public NotesPageData(string filePath)
        {
            this.FilePath = filePath;

            this.VersesNotesPageData = new Dictionary<VersePointer, VerseNotesPageData>();

            if (File.Exists(this.FilePath))
                Deserialize();
            else
                IsNew = true;
        }

        public void Update(ref Microsoft.Office.Interop.OneNote.Application oneNoteApp, 
            VersePointer vp, decimal verseWeight, XmlCursorPosition versePosition,
            bool isChapter, HierarchySearchManager.HierarchyObjectInfo verseHierarchyObjectInfo, Common.HierarchyElementInfo notePageInfo, string notePageContentObjectId,
            bool isImportantVerse, bool force, bool processAsExtendedVerse, out bool rowWasAdded)
        {
            rowWasAdded = true;
        }

        protected void Deserialize()
        {
            
        }

        public void Serialize()
        {
            // не забыть: 
            //  -нумерация
            //  -если ссылок нет, то остальную структуру не надо создавать
            //  -суммируем внутренние веса
        }

        public VerseNotesPageData GetVerseNotesPageData(VersePointer vp)
        {
            if (!VersesNotesPageData.ContainsKey(vp))
            {
                var vnpd = new VerseNotesPageData(vp);
                VersesNotesPageData.Add(vp, vnpd);
                return vnpd;
            }
            else
                return VersesNotesPageData[vp];
        }
    }

    public class VerseNotesPageData : NotesPageLevelBase
    {
        public VersePointer Verse { get; set; }
        
        public Dictionary<string, NotesPageLevel> AllPagesLevels { get; set; }

        public VerseNotesPageData(VersePointer vp)
            : base()
        {
            Verse = vp;
            AllPagesLevels = new Dictionary<string, NotesPageLevel>();
        }

        public override void AddLevel(NotesPageLevel level)
        {
            level.Root = this;
            base.AddLevel(level);
        }
    }


    public enum NotesPageLevelType
    {
        HierarchyElement,
        Page
    }

    public class NotesPageLevel : NotesPageLevelBase
    {
        public string Title { get; set; }        
        public string ID { get; set; }
        public NotesPageLevelType Type { get; set; }

        public NotesPageLevel Parent { get; set; }
        public VerseNotesPageData Root { get; set; }

        private bool _pageLinksWasParsed = false;
        private List<NotesPageLink> _pageLinks;
        public List<NotesPageLink> PageLinks 
        {
            get
            {
                if (!_pageLinksWasParsed)
                {
                    _pageLinks.ForEach(pl => pl.Parse());
                    _pageLinksWasParsed = true;
                }

                return _pageLinks;
            }
        }

        public NotesPageLevel()
            : base()
        {
            _pageLinks = new List<NotesPageLink>();
        }

        public override void AddLevel(NotesPageLevel level)
        {
            level.Root = Root;
            level.Parent = this;

            base.AddLevel(level);

            if (level.Type == NotesPageLevelType.Page)
            {
                if (!Root.AllPagesLevels.ContainsKey(level.ID))
                    Root.AllPagesLevels.Add(level.ID, level);
            }
        }
    }

    public abstract class NotesPageLevelBase
    {
        public Dictionary<string, NotesPageLevel> Levels { get; set; }        

        public virtual void AddLevel(NotesPageLevel level)
        {            
            this.Levels.Add(level.ID, level);
            //todo: sort
        }

        public NotesPageLevelBase()
        {
            Levels = new Dictionary<string, NotesPageLevel>();
        }
    }

    public class NotesPageLink
    {
        public string Href { get; set; }
        public XmlCursorPosition VersePosition { get; set; }
        public decimal VerseWeight { get; set; }        

        internal bool WasParsed { get; set; }

        internal void Parse()
        {
            //Position = ;
            WasParsed = true;
        }
    }
}
