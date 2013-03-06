using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using BibleCommon.Services;
using System.Xml.Linq;
using BibleCommon.Helpers;
using System.Collections.Specialized;
using BibleCommon.Handlers;
using Microsoft.Office.Interop.OneNote;
using BibleCommon.Consts;

namespace BibleCommon.Common
{
    public class NotesPageData
    {
        public bool IsNew { get; set; }
        public string FilePath { get; set; }
        public OrderedDictionary<VersePointer, VerseNotesPageData> VersesNotesPageData { get; set; }                

        public NotesPageData(string filePath)
        {
            this.FilePath = filePath;

            this.VersesNotesPageData = new OrderedDictionary<VersePointer, VerseNotesPageData>();

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

        public void Serialize(ref Application oneNoteApp)
        {
            // не забыть: 
            //  -нумерация
            //  -если ссылок нет, то остальную структуру не надо создавать
            //  -суммируем внутренние веса

            var chapterVersePointer = VersesNotesPageData[0].Verse.GetChapterPointer();
            var chapterLinkHref = OpenBibleVerseHandler.GetCommandUrlStatic(chapterVersePointer, SettingsManager.Instance.ModuleShortName);            
            var pageName = Path.GetFileNameWithoutExtension(this.FilePath);

            var xdoc = new XDocument(
                        new XElement("html", 
                            new XElement("head", 
                                new XElement("title", string.Format("{0} [{1}]", pageName, chapterVersePointer.ChapterName))
                            )));

            var bodyEl = xdoc.Root.AddEl(new XElement("body"));

            bodyEl.Add(new XElement("div", new XAttribute("class", "pageTitle"),
                        new XElement("span",
                            new XAttribute("class", "pageTitleText"),
                            pageName),
                        new XElement("span",
                            new XAttribute("class", "pageTitleLink"),
                            "["),
                        new XElement("a",
                            new XAttribute("class", "pageTitleLink"),
                            new XAttribute("href", chapterLinkHref),
                            chapterVersePointer.ChapterName),
                        new XElement("span",
                            new XAttribute("class", "pageTitleLink"),
                            "]")));

            if (VersesNotesPageData.Count > 1)
            {
                foreach(var verseNotesPageData in VersesNotesPageData.Values)
                {
                    SerializeLevel(ref oneNoteApp, verseNotesPageData, bodyEl, 1, null);
                }                
            }
            else
                SerializeLevel(ref oneNoteApp, VersesNotesPageData[0], bodyEl, 1, null);
        }

        private void SerializeLevel(ref Application oneNoteApp, NotesPageHierarchyLevelBase hierarchyLevel, XElement parentEl, int level, int? index)
        {
            var levelEl = parentEl.AddEl(
                                    new XElement("div", 
                                        new XAttribute("class", "level")));

            if (hierarchyLevel is NotesPagePageLevel)
            {
                GeneratePageLevel(ref oneNoteApp, (NotesPagePageLevel)hierarchyLevel, levelEl, level, index);
            }
            else
            {
                GenerateHierarchyLevel(hierarchyLevel, levelEl, level, index);                

                var childIndex = 0;
                foreach (var childHierarchyLevel in hierarchyLevel.Levels.Values)
                {
                    SerializeLevel(ref oneNoteApp, childHierarchyLevel, levelEl, level + 1, childIndex++);
                }
            }                                    
        }

        private void GeneratePageLevel(ref Application oneNoteApp, NotesPagePageLevel pageLevel, XElement levelEl, int level, int? index)
        {
            var levelTitleEl = levelEl.AddEl(new XElement("p", new XAttribute("class", "pageLevelTitle")));
            levelTitleEl.Add(new XElement("span",
                                new XAttribute("class", "levelTitleIndex"),
                                GetDisplayLevel(level, index)));
            levelEl.Add(new XAttribute("id", pageLevel.ID));

            var levelTitleLinkEl = levelTitleEl.AddEl(new XElement("a", new XAttribute("class", "levelTitleLink"), pageLevel.Title));

            string pageTitleLinkHref;

            if (pageLevel.PageLinks.Count == 1)
            {
                pageTitleLinkHref = pageLevel.PageLinks[0].GetHref(ref oneNoteApp);
                if verseWeight >= Constants.ImportantVerseWeight
            }
            else
            {
                pageTitleLinkHref = OneNoteUtils.GetOrGenerateLinkHref(ref oneNoteApp, null, pageLevel.PageId, pageLevel.PageTitleObjectId, true);                
            }

            levelTitleLinkEl.Add(new XAttribute("href", pageTitleLinkHref));
        }

        private void GenerateHierarchyLevel(NotesPageHierarchyLevelBase hierarchyLevel, XElement levelEl, int level, int? index)
        {
            var levelTitleEl = levelEl.AddEl(new XElement("p", new XAttribute("class", "levelTitle")));

            if (hierarchyLevel is VerseNotesPageData)
            {
                levelTitleEl.Add(
                                new XElement("a",
                                    new XAttribute("class", "verseLink"),
                                    new XAttribute("href", ((VerseNotesPageData)hierarchyLevel).GetVerseLinkHref())
                                ));
            }
            else if (hierarchyLevel is NotesPageHierarchyLevel)
            {
                levelTitleEl.Add(
                                new XElement("span",
                                    new XAttribute("class", "levelTitleIndex"),
                                    GetDisplayLevel(level, index)),
                                new XElement("span",
                                    new XAttribute("class", "levelTitleText"),
                                    ((NotesPageHierarchyLevel)hierarchyLevel).Title));

                levelEl.Add(new XAttribute("id", ((NotesPageHierarchyLevel)hierarchyLevel).ID));
            }
        }        

        private string GetDisplayLevel(int level, int? index)
        {
            throw new NotImplementedException();
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

    public class VerseNotesPageData : NotesPageHierarchyLevelBase, IComparable<VerseNotesPageData>
    {
        public VersePointer Verse { get; set; }

        public Dictionary<string, NotesPagePageLevel> AllPagesLevels { get; set; }

        public string GetVerseLinkHref()
        {
            return OpenBibleVerseHandler.GetCommandUrlStatic(Verse, SettingsManager.Instance.ModuleShortName);
        }

        public VerseNotesPageData(VersePointer vp)
            : base()
        {
            Verse = vp;
            AllPagesLevels = new Dictionary<string, NotesPagePageLevel>();
        }

        public override void AddLevel(NotesPageHierarchyLevel level)
        {
            level.Root = this;
            base.AddLevel(level);
        }

        public int CompareTo(VerseNotesPageData other)
        {
            if (other == null)
                return 1;

            return this.Verse.CompareTo(other.Verse);
        }
    }


    public enum NotesPageLevelType
    {
        HierarchyElement,
        Page
    }

    public class NotesPagePageLevel : NotesPageHierarchyLevel
    {
        public string PageId { get; set; }
        public string PageTitleObjectId { get; set; }

        private bool _pageLinksWasChanged = false;
        private List<NotesPageLink> _pageLinks;        

        public void AddPageLink(NotesPageLink notesPageLink)
        {
            _pageLinksWasChanged = true;
            _pageLinks.Add(notesPageLink);
        }

        public List<NotesPageLink> PageLinks
        {
            get
            {
                if (_pageLinksWasChanged)
                {
                    _pageLinks.ForEach(pl => pl.Parse());
                    _pageLinks = _pageLinks.OrderBy(link => link.VersePosition).ToList();
                    _pageLinksWasChanged = false;
                }

                return _pageLinks;
            }
        }

        public override void AddLevel(NotesPageHierarchyLevel level)
        {
            throw new NotSupportedException();
        }

        public NotesPagePageLevel()
            : base()
        {
            _pageLinks = new List<NotesPageLink>();
        }
    }

    public class NotesPageHierarchyLevel : NotesPageHierarchyLevelBase
    {
        public string Title { get; set; }        
        public string ID { get; set; }        

        public NotesPageHierarchyLevel Parent { get; set; }
        public VerseNotesPageData Root { get; set; }
        

        public NotesPageHierarchyLevel()
            : base()
        {
            
        }

        public override void AddLevel(NotesPageHierarchyLevel level)
        {
            level.Root = Root;
            level.Parent = this;

            base.AddLevel(level);

            if (level is NotesPagePageLevel)
            {
                if (!Root.AllPagesLevels.ContainsKey(level.ID))
                    Root.AllPagesLevels.Add(level.ID, (NotesPagePageLevel)level);
            }
        }
    }

    public abstract class NotesPageHierarchyLevelBase
    {
        public Dictionary<string, NotesPageHierarchyLevel> Levels { get; set; }        

        public virtual void AddLevel(NotesPageHierarchyLevel level)
        {            
            this.Levels.Add(level.ID, level);
            //todo: sort
        }

        public NotesPageHierarchyLevelBase()
        {
            Levels = new Dictionary<string, NotesPageHierarchyLevel>();
        }
    }

    public class NotesPageLink
    {
        private string _href;

        protected bool WasParsed { get; set; }        

        public string PageId { get; set; }
        public string ContentObjectId { get; set; }
        public XmlCursorPosition? VersePosition { get; set; }
        public decimal VerseWeight { get; set; }        

        public string GetHref(ref Application oneNoteApp)
        {
            if (string.IsNullOrEmpty(_href))
            {
                _href = OneNoteUtils.GetOrGenerateLinkHref(ref oneNoteApp, null, PageId, ContentObjectId, true,
                                string.Format("{0}={1}", Constants.QueryParameterKey_VersePosition, VersePosition),
                                string.Format("{0}={1}", Constants.QueryParameterKey_VerseWeight, VerseWeight));
            }

            return _href;
        }

        public void SetHref(string href)
        {
            _href = href;
        }        

        internal void Parse()
        {
            if (!WasParsed)
            {
                if (!string.IsNullOrEmpty(_href))
                {
                    var versePositionString = StringUtils.GetQueryParameterValue(_href, Constants.QueryParameterKey_VersePosition);
                    if (!string.IsNullOrEmpty(versePositionString))
                        VersePosition = new XmlCursorPosition(versePositionString);

                    var verseWeightString = StringUtils.GetQueryParameterValue(_href, Constants.QueryParameterKey_VerseWeight);
                    if (!string.IsNullOrEmpty(verseWeightString))
                        VerseWeight = decimal.Parse(verseWeightString);
                }

                WasParsed = true;
            }
        }
    }
}
