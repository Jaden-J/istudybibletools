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
using System.Xml.XPath;

namespace BibleCommon.Common
{
    public class NotesPageData
    {
        public bool IsNew { get; set; }
        public string FilePath { get; set; }
        public string PageName { get; set; }
        public VersePointer ChapterPoiner { get; set; }

        public Dictionary<VersePointer, VerseNotesPageData> VersesNotesPageData { get; set; }

        public NotesPageData(string filePath, string pageName, VersePointer chapterPointer)
        {
            this.FilePath = filePath;
            this.PageName = pageName;
            this.ChapterPoiner = chapterPointer;

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
            var xdoc = XDocument.Load(this.FilePath);
            foreach (var levelEl in xdoc.Root.Element("body").Elements("div").Where(e => e.Attribute("class").Value == "level"))
            {
                VersePointer vp;
                var verse = levelEl.Element("p").Element("a").Value;
                if (!string.IsNullOrEmpty(verse))
                {
                    var verseNumber = VerseNumber.Parse(verse.Substring(1));
                    vp = new VersePointer(ChapterPoiner, verseNumber.Verse, verseNumber.TopVerse);
                }
                else
                    vp = ChapterPoiner;

                var verseNotesPageData = new VerseNotesPageData(vp);
                VersesNotesPageData.Add(vp, verseNotesPageData);

                AddHierarchySubLevels(levelEl, verseNotesPageData);
            }            
        }

        private void AddHierarchySubLevels(XElement parentLevelEl, NotesPageHierarchyLevelBase parentLevel)
        {
            foreach (var levelEl in parentLevelEl.Elements("div").Where(e => e.Attribute("class").Value == "level"))
            {
                var pEl = levelEl.Element("p");
                if (pEl.Attribute("class").Value == "levelTitle")
                {
                    var title = pEl.Elements("span").Where(e => e.Attribute("class").Value == "levelTitleText").First().Value;
                    var level = new NotesPageHierarchyLevel()
                    {
                        ID = levelEl.Attribute("id").Value,
                        Title = title                        
                    };
                    parentLevel.AddLevel(level);

                    AddHierarchySubLevels(levelEl, level);
                }
                else
                {
                    var aEl = levelEl.Element("p").Element("a");
                    здесь

                }                
            }
        }

        public void Serialize(ref Application oneNoteApp)
        {
            // todo: не забыть: 
            //  -нумерация            

            var chapterVersePointer = VersesNotesPageData.First().Key.GetChapterPointer();
            var chapterLinkHref = OpenBibleVerseHandler.GetCommandUrlStatic(chapterVersePointer, SettingsManager.Instance.ModuleShortName);            

            var xdoc = new XDocument(
                        new XElement("html", 
                            new XElement("head", 
                                new XElement("title", string.Format("{0} {1} [{1}]", PageName,  chapterVersePointer.ChapterName)),
                                new XElement("link", new XAttribute("type", "text/css"), new XAttribute("rel", "stylesheet"), new XAttribute("href", "../../../core.css"))
                            )));

            var bodyEl = xdoc.Root.AddEl(new XElement("body"));

            bodyEl.Add(new XElement("div", new XAttribute("class", "pageTitle"),
                        new XElement("span",
                            new XAttribute("class", "pageTitleText"),
                            PageName),
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
            
            foreach(var verseNotesPageData in VersesNotesPageData.OrderBy(v => v.Key))
            {
                SerializeLevel(ref oneNoteApp, verseNotesPageData.Value, bodyEl, 1, null);
            }

            var folder = Path.GetDirectoryName(this.FilePath);
            if (!Directory.Exists(folder))
                Directory.CreateDirectory(folder);

            xdoc.Save(this.FilePath);
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
                // todo: -если ссылок нет, то остальную структуру не надо создавать - рекурсивно проверять, ести ли хотя бы один NotesPagePageLevel            
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

            var levelTitleLinkEl = levelTitleEl.AddEl(new XElement("a", pageLevel.Title));

            var levelTitleLinkElClass = "levelTitleLink";
            var summarySubLinksWight = pageLevel.PageLinks.Sum(pl => pl.VerseWeight);
            if (summarySubLinksWight >= Constants.ImportantVerseWeight)
                levelTitleLinkElClass += " importantVerseLink";

            string pageTitleLinkHref;

            if (pageLevel.PageLinks.Count == 1)
            {
                var pageLink = pageLevel.PageLinks[0];
                pageTitleLinkHref = pageLink.GetHref(ref oneNoteApp);
                if (!string.IsNullOrEmpty(pageLink.MultiVerseString))
                {
                    levelTitleEl.Add(
                        new XElement("span",
                            new XAttribute("class", "subLinkMultiVerse"),
                            pageLink.MultiVerseString));
                }
            }
            else
            {
                pageTitleLinkHref = OneNoteUtils.GetOrGenerateLinkHref(ref oneNoteApp, null, pageLevel.PageId, pageLevel.PageTitleObjectId, true);                

                var subLinksEl = levelEl.AddEl(new XElement("div", new XAttribute("class", "subLinks")));
                
                var linkIndex = 0;
                foreach (var pageLink in pageLevel.PageLinks)
                {
                    GeneratePageLinkLevel(ref oneNoteApp, pageLink, subLinksEl, linkIndex);
                    linkIndex++;
                }
            }

            levelTitleLinkEl.Add(new XAttribute("class", levelTitleLinkElClass));
            levelTitleLinkEl.Add(new XAttribute("href", pageTitleLinkHref));
        }

        private void GeneratePageLinkLevel(ref Application oneNoteApp, NotesPageLink pageLink, XElement subLinksEl, int linkIndex)
        {
            var importantClassName =
                       pageLink.VerseWeight >= Constants.ImportantVerseWeight
                           ? " importantVerseLink"
                           : string.Empty;

            subLinksEl.Add(
                new XElement("a",
                    new XAttribute("class", "subLink" + importantClassName),
                    new XAttribute("href", pageLink.GetHref(ref oneNoteApp)),
                    string.Format(Resources.Constants.VerseLinkTemplate, linkIndex + 1)));

            if (!string.IsNullOrEmpty(pageLink.MultiVerseString))
            {
                subLinksEl.Add(
                    new XElement("span",
                        new XAttribute("class", "subLinkMultiVerse" + importantClassName),
                        pageLink.MultiVerseString));
            }

            if (linkIndex > 0)
            {
                subLinksEl.Add(
                    new XElement("span",
                        new XAttribute("class", "subLinkDelimeter"),
                        Resources.Constants.VerseLinksDelimiter));
            }
        }        

        private void GenerateHierarchyLevel(NotesPageHierarchyLevelBase hierarchyLevel, XElement levelEl, int level, int? index)
        {
            var levelTitleEl = levelEl.AddEl(new XElement("p", new XAttribute("class", "levelTitle")));

            if (hierarchyLevel is VerseNotesPageData)
            {
                levelTitleEl.Add(
                                new XElement("a",
                                    new XAttribute("class", "verseLink"),
                                    new XAttribute("href", ((VerseNotesPageData)hierarchyLevel).GetVerseLinkHref()),
                                    !((VerseNotesPageData)hierarchyLevel).Verse.IsChapter 
                                        ? ":" + ((VerseNotesPageData)hierarchyLevel).Verse.VerseNumber.ToString()
                                        : string.Empty
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
            return "A.";            
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

        public void AddPageLink(NotesPageLink notesPageLink, VersePointer vp)
        {
            _pageLinksWasChanged = true;
            notesPageLink.MultiVerseString = GetMultiVerseString(vp.ParentVersePointer ?? vp);
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

        private static string GetMultiVerseString(VersePointer vp)
        {
            var result = string.Empty;

            if (vp.IsMultiVerse)
            {
                if (vp.TopChapter != null && vp.TopVerse != null)
                    result = string.Format("({0}:{1}-{2}:{3})", vp.Chapter, vp.Verse, vp.TopChapter, vp.TopVerse);
                else if (vp.TopChapter != null && vp.IsChapter)
                    result = string.Format("({0}-{1})", vp.Chapter, vp.TopChapter);
                else
                    result = string.Format("(:{0}-{1})", vp.Verse, vp.TopVerse);                
            }

            return result;
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
        public string MultiVerseString { get; set; }

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
