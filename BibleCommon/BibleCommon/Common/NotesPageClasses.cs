﻿using System;
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
    public enum NotesPageType
    {
        Verse,
        Chapter,
        Detailed
    }

    public class NotesPageData
    {
        private const string AttributeName_HasValueLinks = "hasValueLinks";
        

        public bool IsNew { get; set; }
        public string FilePath { get; set; }
        public string PageName { get; set; }
        NotesPageType NotesPageType { get; set; }
        public VersePointer ChapterPoiner { get; set; }

        //содержит ли хотя бы одну не подобную ссылку. То есть есть ли на эту страницу уже ссылка "Заметки" в записной книжке "Библия".
        public bool HasValueLinks { get; set; }

        public Dictionary<VerseNumber, VerseNotesPageData> VersesNotesPageData { get; set; }

        public NotesPageData(string filePath, string pageName, NotesPageType notesPageType, VersePointer chapterPointer, bool toDeserializeIfExists)
        {
            this.FilePath = filePath;
            this.PageName = pageName;
            this.NotesPageType = notesPageType;
            this.ChapterPoiner = chapterPointer;

            this.VersesNotesPageData = new Dictionary<VerseNumber, VerseNotesPageData>();

            if (toDeserializeIfExists && File.Exists(this.FilePath))
            {
                Deserialize();
                IsNew = !HasValueLinks;
            }
            else
                IsNew = true;
        }      

        protected void Deserialize()
        {            
            var xdoc = XDocument.Load(this.FilePath);
            var bodyEl = xdoc.Root.Element("body");

            var hvlAttribute = bodyEl.Attribute(AttributeName_HasValueLinks);
            if (hvlAttribute != null)
                this.HasValueLinks = Convert.ToBoolean(hvlAttribute.Value);

            foreach (var levelEl in bodyEl.GetByClass("div", "verseLevel"))
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

                var verseNotesPageData = new VerseNotesPageData(vp, this);
                VersesNotesPageData.Add(vp.VerseNumber, verseNotesPageData);

                AddHierarchySubLevels(levelEl, verseNotesPageData);
            }            
        }

        private void AddHierarchySubLevels(XElement parentLevelEl, NotesPageHierarchyLevelBase parentLevel)
        {
            foreach (var levelEl in parentLevelEl.Elements("ol").Elements("li"))
            {
                if (levelEl.Attribute("class").Value.Contains("pageLevel"))
                {
                    AddHierarchyPageSubLevel(levelEl, parentLevel); 
                }
                else if (levelEl.HasElements)  // а то уже нескольким людям поставил версию, в которой иногда могли создаваться пустые ноды
                {
                    AddHierarchyLevelSubLevel(levelEl, parentLevel);
                }               
            }
        }

        private void AddHierarchyPageSubLevel(XElement levelEl, NotesPageHierarchyLevelBase parentLevel)
        {
            var aEl = levelEl.Element("p").Element("a");

            var pageLevel = new NotesPagePageLevel()
            {
                Id = levelEl.Attribute(Constants.NotesPageElementAttributeName_SyncId).Value,
                Title = aEl.Value                
            };
            var subLinksEl = levelEl.Element("table");
            if (subLinksEl != null)
            {
                subLinksEl = subLinksEl.Element("tr");
                pageLevel.SetPageTitleLinkHref(aEl.Attribute("href").Value);

                foreach (var subLinkEl in subLinksEl.GetByClass("td", "subLink"))
                {
                    pageLevel.AddPageLink(GetNotesPageLink(subLinkEl.Element("a"), subLinkEl, "td"));                    
                }
            }
            else
            {
                pageLevel.AddPageLink(GetNotesPageLink(aEl, aEl, "span"));                
            }

            Application oneNoteApp = null;
            parentLevel.AddLevel(ref oneNoteApp, pageLevel, null, false);
        }

        private NotesPageLink GetNotesPageLink(XElement aEl, XElement linkEl, string multiVerseTagName)
        {
            var href = aEl.Attribute("href").Value;

            var pageLink = new NotesPageLink(href);
            
            var multiVerseEl = (XElement)linkEl.NextNode;
            if (multiVerseEl != null && multiVerseEl.Name == multiVerseTagName && multiVerseEl.HasClass("subLinkMultiVerse"))
                pageLink.MultiVerseString = multiVerseEl.Value;

            return pageLink;
        }

        private void AddHierarchyLevelSubLevel(XElement levelEl, NotesPageHierarchyLevelBase parentLevel)
        {
            var title = levelEl.Element("p").GetByClass("span", "levelTitleText").First().Value;            
            var level = new NotesPageHierarchyLevel()
            {
                Id = levelEl.Attribute(Constants.NotesPageElementAttributeName_SyncId).Value,
                HierarchyType = GetHierarchyTypeFromClassName(levelEl.Attribute("class").Value),
                Title = title,
                Collapsed = levelEl.Element("ol").Attribute("class").Value.Contains("collapsedLevel")
            };

            Application oneNoteApp = null;
            parentLevel.AddLevel(ref oneNoteApp, level, null, false);

            AddHierarchySubLevels(levelEl, level);
        }        

        public void Serialize(ref Application oneNoteApp, AnalyzedVersesService analyzedVersesService, bool writeNotesPage = true)
        {            
            var chapterVersePointer = VersesNotesPageData.First().Value.Verse.GetChapterPointer();            

            var xdoc = new XDocument(
                        new XElement("html", 
                            new XElement("head", 
                                new XElement("meta", new XAttribute("http-equiv", "X-UA-Compatible"), new XAttribute("content", "IE=edge")),
                                new XElement("title", string.Format("{0} [{1}]", PageName,  chapterVersePointer.ChapterName)),
                                new XElement("link", new XAttribute("type", "text/css"), new XAttribute("rel", "stylesheet"), new XAttribute("href", "../../../" + Constants.NotesPageStyleFileName)),
                                new XElement("link", new XAttribute("type", "text/css"), new XAttribute("rel", "stylesheet"), new XAttribute("href", "../../../" + Constants.NotesPageTrackbarStyleFileName)),                                
                                new XElement("script", new XAttribute("type", "text/javascript"), new XAttribute("src", "../../../" + Constants.NotesPageJQueryScriptFileName), ";"),
                                new XElement("script", new XAttribute("type", "text/javascript"), new XAttribute("src", "../../../" + Constants.NotesPageScriptFileName), ";"),
                                new XElement("script", new XAttribute("type", "text/javascript"), new XAttribute("src", "../../../" + Constants.NotesPageTrackbarScriptFileName), ";")
                            )));

            var bodyEl = xdoc.Root.AddEl(new XElement("body", new XAttribute(AttributeName_HasValueLinks, HasValueLinks)));

            AddPageTitle(bodyEl, chapterVersePointer);           
            
            foreach(var verseNotesPageData in VersesNotesPageData.OrderBy(v => v.Key))
            {
                SerializeLevel(ref oneNoteApp, verseNotesPageData.Value, bodyEl, 0, null, analyzedVersesService);
            }

            var folder = Path.GetDirectoryName(this.FilePath);
            if (!Directory.Exists(folder))
                Directory.CreateDirectory(folder);

            xdoc.AddFirst(new XDocumentType("html", null, null, null));

            if (writeNotesPage)
                xdoc.Save(this.FilePath);
        }

        public VerseNotesPageData GetVerseNotesPageData(VersePointer vp)
        {
            if (!VersesNotesPageData.ContainsKey(vp.VerseNumber))
            {
                var vnpd = new VerseNotesPageData(vp, this);
                VersesNotesPageData.Add(vp.VerseNumber, vnpd);

                return vnpd;
            }
            else
                return VersesNotesPageData[vp.VerseNumber];
        }

        private void AddPageTitle(XElement bodyEl, VersePointer chapterVersePointer)
        {            
            var chapterLinkHref = OpenBibleVerseHandler.GetCommandUrlStatic(chapterVersePointer, SettingsManager.Instance.ModuleShortName);         

            bodyEl.Add(new XElement("table", new XAttribute("class", "pageTitleText"), new XAttribute("cellpadding", "0"), new XAttribute("cellspacing", "0"),
                            new XElement("tr",
                                new XElement("td",
                                    new XAttribute("class", "pageTitleText"),
                                    PageName),
                                new XElement("td",
                                    new XAttribute("class", "pageTitleLink"),
                                    "["),
                                new XElement("td",
                                    new XAttribute("class", "pageTitleLink"),
                                    new XElement("a",
                                        new XAttribute("class", "pageTitleLink"),
                                        new XAttribute("href", chapterLinkHref),
                                        chapterVersePointer.ChapterName)),
                                new XElement("td",
                                    new XAttribute("class", "pageTitleLink"),
                                    "]"))));


            XObject pageTitleAdditionalLinkEl;
            if (NotesPageType != Common.NotesPageType.Chapter)
            {
                pageTitleAdditionalLinkEl = new XElement("a",
                                                    new XAttribute("class", "chapterNotesPage"),
                                                    new XAttribute("href",
                                                        OpenNotesPageHandler.GetCommandUrlStatic(chapterVersePointer, SettingsManager.Instance.ModuleShortName, Common.NotesPageType.Chapter)),
                                                    Resources.Constants.ChapterNotes);
            }
            else
                pageTitleAdditionalLinkEl = new XElement("br");

            bodyEl.Add(new XElement("table", new XAttribute("class", "pageTitleRightBlock"), new XAttribute("cellpadding", "0"), new XAttribute("cellspacing", "0"),
                            new XElement("tr",
                                new XElement("td", new XAttribute("class", "chapterNotesPage"),
                                     pageTitleAdditionalLinkEl)),
                            new XElement("tr",
                                GenerateFilterBlock())));                                
        }

        private XElement GenerateFilterBlock()
        {
            return new XElement("td", new XAttribute("class", "filter"),
                        new XElement("div", new XAttribute("class", "filterLink"),
                            new XElement("a", new XAttribute("id", "filterLink"), new XAttribute("class", "filterLink"), new XAttribute("href", "#"), Resources.Constants.Filter)),
                        new XElement("div", new XAttribute("id", "filter"), new XAttribute("class", "filterPopup"), new XAttribute("style", "display:none;"),
                            new XElement("div", new XAttribute("class", "filterPopup_background"),
                                new XElement("div", new XAttribute("class", "filterPopup_controls"),
                                    new XElement("div", new XAttribute("class", "filterPopup_header"),
                                        new XElement("div", new XAttribute("class", "filterPopup_section notebooksFilter"),
                                            new XElement("div", new XAttribute("class", "filterTitle"), Resources.Constants.FilterPopupNotebooks),
                                            new XElement("div", new XAttribute("class", "filterPopup_label"),
                                                new XElement("table", new XAttribute("id", "notebooksFilterTable"), new XAttribute("class", "notebooksFilter"), 
                                                    new XAttribute("cellpadding", "0"), new XAttribute("cellspacing", "0"), " "))),
                                        new XElement("div", new XAttribute("class", "filterPopup_hr filterPopup_hrheight notebooksFilter"), " "),
                                        new XElement("div", new XAttribute("class", "filterPopup_section"),
                                            new XElement("span", new XAttribute("class", "filterTitle"), Resources.Constants.FilterPopupTrackbar),
                                            new XElement("div", new XAttribute("class", "filterPopup_textsection"),
                                                new XElement("span", new XAttribute("id", "filterVerseWeightTitle"), " "),
                                                new XElement("span", new XAttribute("id", "filterVerseWeightDescription"), Resources.Constants.FilterPopupAndMore),
                                                new XElement("div", new XAttribute("id", "filterVerseWeight"), " "))),
                                        new XElement("div", new XAttribute("class", "filterPopup_hr"), " "),
                                        new XElement("div", new XAttribute("class", "filterPopup_textsection detailedNotes"),
                                            new XElement("input", new XAttribute("type", "checkbox"), new XAttribute("id", "chkDetailedNotes"), new XAttribute("class", "detailedNotes")),
                                            new XElement("label", new XAttribute("for", "chkDetailedNotes"), new XAttribute("class", "detailedNotes filterTitle"), 
                                                new XAttribute("title", Resources.Constants.DetailedNotes), Resources.Constants.DetailedNotes)),
                                        new XElement("div", new XAttribute("class", "filterPopup_hr filterPopup_hrheight saveFilterSettings"), " "),
                                        new XElement("div", new XAttribute("class", "filterPopup_textsection saveFilterSettingsLink saveFilterSettings"),
                                            new XElement("a", new XAttribute("id", "saveFilterSettingsLink"), 
                                                new XAttribute("class", "filterTitle"), new XAttribute("href", Constants.NoLinkTransmitHref), 
                                                new XAttribute("title", Resources.Constants.FilterPopupSave), Resources.Constants.FilterPopupSave)))))));
        }

        private void SerializeLevel(ref Application oneNoteApp, NotesPageHierarchyLevelBase hierarchyLevel, XElement parentEl, int levelIndex, int? index, AnalyzedVersesService analyzedVersesService)
        {
            var levelEl = parentEl.AddEl(
                                    new XElement(hierarchyLevel is VerseNotesPageData ? "div" : "li", 
                                        new XAttribute("class", GetLevelClassName(hierarchyLevel, levelIndex))));                       

            if (hierarchyLevel is NotesPagePageLevel)
            {
                GeneratePageLevel(ref oneNoteApp, (NotesPagePageLevel)hierarchyLevel, levelEl, levelIndex, index, analyzedVersesService);
            }
            else
            {
                if (HierarchyLevelContainsPageLinks(hierarchyLevel))
                {
                    GenerateHierarchyLevel(hierarchyLevel, levelEl, levelIndex, index);

                    if (hierarchyLevel.Levels.Count > 0)  // хотя по идеи не может у него не быть детей то...
                    {
                        var className = "levelChilds level" + (levelIndex + 1);
                        if (hierarchyLevel is NotesPageHierarchyLevel && ((NotesPageHierarchyLevel)hierarchyLevel).Collapsed)
                            className += " collapsedLevel";
                        levelEl = levelEl.AddEl(
                                            new XElement("ol",
                                                new XAttribute("class", className)));

                        var childIndex = 0;
                        foreach (var childHierarchyLevel in hierarchyLevel.Levels.Values)
                        {
                            SerializeLevel(ref oneNoteApp, childHierarchyLevel, levelEl, levelIndex + 1, childIndex++, analyzedVersesService);
                        }
                    }
                }
                else
                    levelEl.Remove();
            }                                    
        }       

        private HierarchyElementType GetHierarchyTypeFromClassName(string className)
        {
            if (className.Contains("notebookLevel"))
                return HierarchyElementType.Notebook;
            else if (className.Contains("sectionGroupLevel"))
                return HierarchyElementType.SectionGroup;
            else if (className.Contains("sectionLevel"))
                return HierarchyElementType.Section;
            else if (className.Contains("pageLevel"))
                return HierarchyElementType.Page;
            else
                throw new NotSupportedException(className);
        }  

        private object GetLevelClassName(NotesPageHierarchyLevelBase hierarchyLevel, int levelIndex)
        {
            if (hierarchyLevel is VerseNotesPageData)
                return "verseLevel";

            switch (((NotesPageHierarchyLevel)hierarchyLevel).HierarchyType)
            {
                case HierarchyElementType.Notebook:
                    return "notebookLevel level" + levelIndex;
                case HierarchyElementType.SectionGroup:
                    return "sectionGroupLevel level" + levelIndex;
                case HierarchyElementType.Section:
                    return "sectionLevel level" + levelIndex;
                case HierarchyElementType.Page:
                    return "pageLevel level" + levelIndex;
                default:
                    throw new NotSupportedException(((NotesPageHierarchyLevel)hierarchyLevel).HierarchyType.ToString());
            }
        }       

        private bool HierarchyLevelContainsPageLinks(NotesPageHierarchyLevelBase hierarchyLevel)
        {
            return hierarchyLevel.Levels.Values.Any(hl => hl is NotesPagePageLevel || HierarchyLevelContainsPageLinks(hl));
        }

        private void GeneratePageLevel(ref Application oneNoteApp, NotesPagePageLevel pageLevel, XElement levelEl, int levelIndex, int? index, AnalyzedVersesService analyzedVersesService)
        {
            levelEl.Add(new XAttribute(Constants.NotesPageElementAttributeName_SyncId, pageLevel.Id));            

            var levelTitleEl = levelEl.AddEl(
                                        new XElement("p", 
                                            new XAttribute("class", "pageLevelTitle")));                        

            var levelTitleLinkEl = levelTitleEl.AddEl(new XElement("a", pageLevel.Title));

            var levelTitleLinkElClass = "levelTitleLink";
            var summarySubLinksOverallWight = pageLevel.PageLinks.Sum(pl => pl.VerseWeight);
            if (summarySubLinksOverallWight >= Constants.ImportantVerseWeight)
                levelTitleLinkElClass += " importantVerseLink";

            var notDetailedPageLinks = pageLevel.PageLinks.Where(pl => !pl.IsDetailed).ToList();
            var summarySubLinksNonDetailedWight = notDetailedPageLinks.Sum(pl => pl.VerseWeight);

            analyzedVersesService.UpdateVerseInfo(pageLevel.Root.Verse, 
                summarySubLinksNonDetailedWight, summarySubLinksOverallWight, 
                notDetailedPageLinks.Count == 0, GetNotebookName(pageLevel));

            var hiddenAllPageLevel = false;

            if (pageLevel.PageLinks.Count == 1)
            {
                var pageLink = pageLevel.PageLinks[0];
                pageLevel.SetPageTitleLinkHref(pageLink.GetHref(ref oneNoteApp));

                if (pageLink.IsDetailed)
                    hiddenAllPageLevel = true;                    

                if (!string.IsNullOrEmpty(pageLink.MultiVerseString))
                {
                    levelTitleEl.AddEl(
                        new XElement("span",
                            new XAttribute("class", "subLinkMultiVerse"),
                            pageLink.MultiVerseString));
                }
            }
            else
            {
                var subLinksEl = levelEl.AddEl(new XElement("table", 
                                                new XAttribute("class", "subLinks"),
                                                new XAttribute("cellpadding", "0"),
                                                new XAttribute("cellspacing", "0")));
                subLinksEl = subLinksEl.AddEl(new XElement("tr"));
                
                var linkIndex = 0;                
                foreach (var pageLink in pageLevel.PageLinks)
                {
                    GeneratePageLinkLevel(ref oneNoteApp, pageLink, subLinksEl, linkIndex, linkIndex == pageLevel.PageLinks.Count - 1);
                    linkIndex++;
                }

                if (notDetailedPageLinks.Count == 0)
                    hiddenAllPageLevel = true;
            }

            if (hiddenAllPageLevel)
            {
                levelEl.SetAttributeValue("class", levelEl.Attribute("class").Value + " detailed");
                levelTitleLinkElClass += " detailed";
            }

            levelTitleLinkEl.Add(new XAttribute("class", levelTitleLinkElClass));
            levelTitleLinkEl.Add(new XAttribute("href", pageLevel.GetPageTitleLinkHref(ref oneNoteApp)));
        }

        private string GetNotebookName(NotesPageHierarchyLevelBase level)
        {
            if (level == null || level is VerseNotesPageData)
                return null;

            var hierarchyLevel = (NotesPageHierarchyLevel)level;
            if (hierarchyLevel.HierarchyType == HierarchyElementType.Notebook)
                return hierarchyLevel.Id;
            else
                return GetNotebookName(hierarchyLevel.Parent);
        }

        private void GeneratePageLinkLevel(ref Application oneNoteApp, NotesPageLink pageLink, XElement subLinksEl, int linkIndex, bool isLast)
        {
            var detailedClassName = pageLink.IsDetailed ? " detailed" : string.Empty;                                    

            var importantClassName =
                       pageLink.VerseWeight >= Constants.ImportantVerseWeight
                           ? " importantVerseLink"
                           : string.Empty;

            subLinksEl.Add(new XElement("td", new XAttribute("class", "subLink" + detailedClassName),
                                new XElement("a",
                                    new XAttribute("class", "subLinkLink" + importantClassName + detailedClassName),
                                    new XAttribute("href", pageLink.GetHref(ref oneNoteApp)),
                                    string.Format(Resources.Constants.VerseLinkTemplate, linkIndex + 1))));         

            if (!string.IsNullOrEmpty(pageLink.MultiVerseString))
            {
                subLinksEl.Add(
                    new XElement("td",
                        new XAttribute("class", "subLinkMultiVerse" + importantClassName + detailedClassName),
                        pageLink.MultiVerseString));
            }

            if (!isLast)
            {
                subLinksEl.Add(
                    new XElement("td",
                        new XAttribute("class", "subLinkDelimeter" + detailedClassName),
                        Resources.Constants.VerseLinksDelimiter));
            }
        }

        private void GenerateHierarchyLevel(NotesPageHierarchyLevelBase hierarchyLevel, XElement levelEl, int level, int? index)
        {
            AddImgCollapse(hierarchyLevel, levelEl);                    

            var levelTitleEl = levelEl.AddEl(new XElement("p", new XAttribute("class", "levelTitle")));

            if (hierarchyLevel is VerseNotesPageData)
            {
                levelTitleEl.Add(
                                new XElement("a",
                                    new XAttribute("class", "verseLink"),
                                    new XAttribute("id", ((VerseNotesPageData)hierarchyLevel).Verse.Verse.GetValueOrDefault(0)),
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
                                    new XAttribute("class", "levelTitleText"),
                                    ((NotesPageHierarchyLevel)hierarchyLevel).Title));

                levelEl.Add(new XAttribute(Constants.NotesPageElementAttributeName_SyncId, ((NotesPageHierarchyLevel)hierarchyLevel).Id));                
            }
        }

        private void AddImgCollapse(NotesPageHierarchyLevelBase hierarchyLevel, XElement levelEl)
        {
            var className = hierarchyLevel is VerseNotesPageData ? "verseLevel minus" : "minus";
            var imgFileName = "none.png";
            if (hierarchyLevel is NotesPageHierarchyLevel && ((NotesPageHierarchyLevel)hierarchyLevel).Collapsed)
            {
                className += " collapsed";
                imgFileName = "plus.png";
            }

            levelEl.Add(new XElement("img", new XAttribute("class", className), new XAttribute("src", "../../../images/" + imgFileName)));
        }

        private string GetDisplayLevel(int level, int? index)
        {
            return "A.";            
        }               
    }

    public class VerseNotesPageData : NotesPageHierarchyLevelBase, IComparable<VerseNotesPageData>
    {
        public NotesPageData Parent { get; set; }
        public VersePointer Verse { get; set; }        

        public Dictionary<string, NotesPagePageLevel> AllPagesLevels { get; set; }

        public string GetVerseLinkHref()
        {
            return OpenBibleVerseHandler.GetCommandUrlStatic(Verse, SettingsManager.Instance.ModuleShortName);
        }

        public VerseNotesPageData(VersePointer vp, NotesPageData parent)
            : base()
        {
            Verse = vp;
            Parent = parent;
            AllPagesLevels = new Dictionary<string, NotesPagePageLevel>();
        }

        public override void AddLevel(ref Application oneNoteApp, NotesPageHierarchyLevel level, HierarchyElementInfo notePageHierarchyElInfo, bool toSort)
        {
            level.Root = this;
            level.Parent = this;
            base.AddLevel(ref oneNoteApp, level, notePageHierarchyElInfo, toSort);
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

        private string _pageTitleLinkHref;
        public string GetPageTitleLinkHref(ref Application oneNoteApp)
        {
            if (string.IsNullOrEmpty(_pageTitleLinkHref) && !string.IsNullOrEmpty(PageId))            
                _pageTitleLinkHref = OneNoteUtils.GetOrGenerateLinkHref(ref oneNoteApp, null, PageId, PageTitleObjectId, true);

            return _pageTitleLinkHref;
        }

        public void SetPageTitleLinkHref(string pageId, string pageTitleObjectId)
        {
            this.PageId = pageId;
            this.PageTitleObjectId = pageTitleObjectId;
        }

        public void SetPageTitleLinkHref(string href)
        {
            _pageTitleLinkHref = href;
        }

        public void AddPageLink(NotesPageLink notesPageLink, VersePointer vp)
        {
            _pageLinksWasChanged = true;
            notesPageLink.MultiVerseString = GetMultiVerseString(vp.ParentVersePointer ?? vp);
            _pageLinks.Add(notesPageLink);
            notesPageLink.Parent = this;
            if (!notesPageLink.IsDetailed && !Root.Parent.HasValueLinks)
                Root.Parent.HasValueLinks = true;
        }

        internal void AddPageLink(NotesPageLink notesPageLink)
        {
            _pageLinksWasChanged = true;            
            _pageLinks.Add(notesPageLink);
            notesPageLink.Parent = this;
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

        public override void AddLevel(ref Application oneNoteApp, NotesPageHierarchyLevel level, HierarchyElementInfo notePageHierarchyElInfo, bool toSort)
        {
            throw new NotSupportedException();
        }

        public NotesPagePageLevel()
            : base()
        {
            _pageLinks = new List<NotesPageLink>();
            this.HierarchyType = HierarchyElementType.Page;
        }

        private static string GetMultiVerseString(VersePointer vp)
        {
            var result = string.Empty;

            if (vp.IsMultiVerse)
                result = string.Format("({0})", vp.GetLightMultiVerseString());            

            return result;
        }
    }

    public class NotesPageHierarchyLevel : NotesPageHierarchyLevelBase
    {
        public string Title { get; set; }        
        public string Id { get; set; }

        public HierarchyElementType HierarchyType { get; set; }
        public string OneNoteId { get; set; }   // идентификатор элемента в XML OneNote. Свойство Id же может содержать всё, что угодно (например, UniqueName - то есть имя раздела, записной книжки...)

        public NotesPageHierarchyLevelBase Parent { get; set; }
        public VerseNotesPageData Root { get; set; }
        public bool Collapsed { get; set; }

        public NotesPageHierarchyLevel()
            : base()
        {
            
        }

        public override void AddLevel(ref Application oneNoteApp, NotesPageHierarchyLevel level, HierarchyElementInfo notePageHierarchyElInfo, bool toSort)
        {
            level.Root = Root;
            level.Parent = this;

            base.AddLevel(ref oneNoteApp, level, notePageHierarchyElInfo, toSort);

            if (level is NotesPagePageLevel)
            {
                if (!Root.AllPagesLevels.ContainsKey(level.Id))
                    Root.AllPagesLevels.Add(level.Id, (NotesPagePageLevel)level);
            }
        }

        public void TryToSortInParent(ref Application oneNoteApp, HierarchyElementInfo hierarchyElementInfo)
        {
            bool levelWasFound;
            var prevLevelIndex = GetPrevLevelIndex(ref oneNoteApp, this, hierarchyElementInfo, out levelWasFound);
            var currentLevelIndex = this.Parent.Levels.IndexOf(this.Id);
            var newLevelIndex = prevLevelIndex + 1;

            var needToMoveOrInsert = !(levelWasFound && newLevelIndex == 0);  // иначе ссылка стоит в начале и она должана там стоять
            if (needToMoveOrInsert && levelWasFound)
            {
                if (newLevelIndex == currentLevelIndex)  // ссылка и так уже на правильном месте
                    needToMoveOrInsert = false;
            }

            if (needToMoveOrInsert)  
            {
                this.Parent.Levels.RemoveAt(currentLevelIndex);
                if (newLevelIndex > currentLevelIndex) newLevelIndex--;
                this.Parent.Levels.Insert(newLevelIndex, this.Id, this);
            }
        }
    }

    public abstract class NotesPageHierarchyLevelBase
    {
        public OrderedDictionary<string, NotesPageHierarchyLevel> Levels { get; set; }        

        public virtual void AddLevel(ref Application oneNoteApp, NotesPageHierarchyLevel level, HierarchyElementInfo notePageHierarchyElInfo, bool toSort)
        {
            if (toSort && oneNoteApp != null && level != null)
            {
                bool levelWasFound;
                var prevLevelIndex = GetPrevLevelIndex(ref oneNoteApp, level, notePageHierarchyElInfo, out levelWasFound);                

                this.Levels.Insert(prevLevelIndex + 1, level.Id, level);
            }
            else
                this.Levels.Add(level.Id, level);
        }

        protected int GetPrevLevelIndex(ref Application oneNoteApp, NotesPageHierarchyLevel level, HierarchyElementInfo notePageHierarchyElInfo, out bool levelWasFound)
        {
            var index = -1;
            levelWasFound = false;

            try
            {                
                XElement parentHierarchy, noteLinkInHierarchy = null;

                var notebookHierarchy = ApplicationCache.Instance.GetHierarchy(ref oneNoteApp, notePageHierarchyElInfo.NotebookId, HierarchyScope.hsPages);  //from cache            
                if (level.HierarchyType != HierarchyElementType.Notebook)
                {
                    noteLinkInHierarchy = notebookHierarchy.Content.Root.XPathSelectElement(
                        string.Format("//one:{0}[@ID=\"{1}\"]", HierarchyElementInfo.GetElementName(level.HierarchyType), notePageHierarchyElInfo.Id),
                        notebookHierarchy.Xnm);

                    parentHierarchy = noteLinkInHierarchy.Parent;
                }
                else
                    parentHierarchy = ApplicationCache.Instance.GetHierarchy(ref oneNoteApp, null, HierarchyScope.hsNotebooks).Content.Root;

                if (noteLinkInHierarchy == null)
                    noteLinkInHierarchy = parentHierarchy.XPathSelectElement(string.Format("*[@ID=\"{0}\"]", notePageHierarchyElInfo.Id), notebookHierarchy.Xnm);

                var prevNodesInHierarchy = noteLinkInHierarchy.NodesBeforeSelf();

                if (prevNodesInHierarchy.Count() != 0)
                {
                    foreach (var otherLevel in level.Parent.Levels.Values)
                    {
                        if (!string.IsNullOrEmpty(notePageHierarchyElInfo.ManualId))
                        {
                            if (otherLevel.Title == notePageHierarchyElInfo.UniqueTitle)
                            {
                                levelWasFound = true;
                                break;
                            }
                            else
                            {
                                if (notePageHierarchyElInfo.UniqueTitle.CompareTo(otherLevel.Title) < 0)
                                    break;
                            }
                        }
                        else
                        {
                            if (otherLevel.Id == notePageHierarchyElInfo.UniqueName)
                            {
                                levelWasFound = true;
                                break;
                            }
                            else
                            {
                                XElement existingLinkInHierarchy;

                                if (notePageHierarchyElInfo.Type == HierarchyElementType.Page)
                                {
                                    existingLinkInHierarchy = parentHierarchy.XPathSelectElement(
                                                string.Format("one:Page[./one:Meta[@name=\"{0}\" and @content=\"{1}\"]]", Constants.Key_SyncId, otherLevel.Id),
                                                notebookHierarchy.Xnm);

                                    if (existingLinkInHierarchy == null)  // ещё не обновили метаданные в кэше                                
                                        existingLinkInHierarchy = parentHierarchy.XPathSelectElement(string.Format("one:Page[@ID='{0}']", otherLevel.Id), notebookHierarchy.Xnm);
                                }
                                else
                                {
                                    existingLinkInHierarchy = parentHierarchy.XPathSelectElement(
                                                string.Format("*[@name=\"{0}\"]", otherLevel.Id),
                                                notebookHierarchy.Xnm);
                                }

                                if (!prevNodesInHierarchy.Contains(existingLinkInHierarchy))
                                    break;
                            }
                        }
                        index++;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogError(ex);
            }

            return index;
        }

        public NotesPageHierarchyLevelBase()
        {
            Levels = new OrderedDictionary<string, NotesPageHierarchyLevel>();
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
        public VersePointer VersePointer { get; set; }
        public NotesPagePageLevel Parent { get; set; }

        private bool _isDetailed;
        public bool IsDetailed 
        {
            get
            {
                return _isDetailed;
            }
            set
            {
                _isDetailed = value;

                if (Parent != null && !_isDetailed && !Parent.Root.Parent.HasValueLinks)
                {
                    Parent.Root.Parent.HasValueLinks = true;
                }
            }
        }

        public string GetHref(ref Application oneNoteApp)
        {
            if (string.IsNullOrEmpty(_href))
            {
                _href = OneNoteUtils.GetOrGenerateLinkHref(ref oneNoteApp, null, PageId, ContentObjectId, true,
                                string.Format("{0}={1}", Constants.QueryParameterKey_VersePosition, VersePosition),
                                string.Format("{0}={1}", Constants.QueryParameterKey_VerseWeight, VerseWeight),
                                string.Format("{0}={1}", Constants.QueryParameterKey_IsDetailedLink, IsDetailed));
            }

            return _href;
        }

        public NotesPageLink()
        {
        }

        public NotesPageLink(string href)
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

                    var isDetailedLink = StringUtils.GetQueryParameterValue(_href, Constants.QueryParameterKey_IsDetailedLink);
                    if (!string.IsNullOrEmpty(isDetailedLink))
                        IsDetailed = bool.Parse(isDetailedLink);
                }

                WasParsed = true;
            }
        }
    }
}
