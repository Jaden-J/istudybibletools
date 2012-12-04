using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Microsoft.Office.Interop.OneNote;
using System.Xml;
using System.Xml.XPath;
using BibleCommon.Common;
using BibleCommon.Helpers;


namespace BibleCommon.Services
{
    public static class HierarchySearchManager
    {
        public enum HierarchyStage
        {
            SectionGroup,
            Section,
            Page,
            ContentPlaceholder
        }

        public class VerseObjectInfo
        {            
            public string ObjectId { get; set; }
            public VerseNumber? VerseNumber { get; set; } // Мы, например, искали Быт 4:4 (модуль IBS). А нам вернули Быт 4:3. Здесь будем хранить "3-4".
            public string ObjectHref { get; set; }

            public VerseObjectInfo()
            {
            }

            public VerseObjectInfo(VersePointerLink link)
            {
                this.ObjectId = link.ObjectId;
                this.VerseNumber = link.VerseNumber;
                this.ObjectHref = link.Href;
            }

            public bool IsVerse { get { return VerseNumber != null; } }
        }

        [Serializable]
        public class HierarchyObjectInfo
        {            
            public string SectionId { get; set; }
            public string PageId { get; set; }
            public string PageName { get; set; }
            public VerseObjectInfo VerseInfo { get; set; }
            public Dictionary<VersePointer, VerseObjectInfo> AdditionalObjectsIds { get; set; }                        
            public List<VerseObjectInfo> GetAllObjectsIds()
            {
                var result = new List<VerseObjectInfo>();

                if (VerseInfo != null)
                    result.Add(VerseInfo);

                result.AddRange(AdditionalObjectsIds.Values);

                return result;
            }

            public VerseNumber? VerseNumber
            {
                get
                {
                    if (VerseInfo != null)
                        return VerseInfo.VerseNumber;

                    return null;
                }
            }

            public string VerseContentObjectId
            {
                get
                {
                    if (VerseInfo != null)
                        return VerseInfo.ObjectId;

                    return null;
                }
            }

            public HierarchyObjectInfo()
            {
                this.AdditionalObjectsIds = new Dictionary<VersePointer, VerseObjectInfo>();
            }
        }

        public enum HierarchySearchResultType
        {
            NotFound,            
            PartlyFound,  // например надо было найти стих, а нашли только страницу (если искали Быт 1:120)
            Successfully,            
        }

        public enum FindVerseLevel
        {
            OnlyFirstVerse,
            OnlyVersesOfFirstChapter,    // пока не работает, если указана ссылка, включающая в себя несколько глав (например, 5:6-6:7)
            AllVerses
        }
        
        [Serializable]
        public class HierarchySearchResult
        {
            public HierarchyObjectInfo HierarchyObjectInfo { get; set; } // дополнительная информация о найденном объекте            
            public HierarchyStage HierarchyStage { get; set; }
            public HierarchySearchResultType ResultType { get; set; }     

            public HierarchySearchResult()
            {
                HierarchyObjectInfo = new HierarchyObjectInfo();
            }


            /// <summary>
            /// 
            /// </summary>
            /// <param name="oneNoteApp"></param>
            /// <param name="vp"></param>
            /// <param name="versePointerLink"></param>
            /// <param name="findAllVerseObjects">Поиск осуществляется только по текущей главе. Чтобы найти все стихи во всех главах (если например ссылка 3:4-6:8), то надо отдельно вызвать GetAllIncludedVersesExceptFirst</param>
            public HierarchySearchResult(ref Application oneNoteApp, string bibleNotebookId, VersePointer vp, VersePointerLink versePointerLink, FindVerseLevel findAllVerseObjects)
                : this()
            {                
                this.ResultType = vp.IsChapter == versePointerLink.IsChapter ? HierarchySearchResultType.Successfully : HierarchySearchResultType.PartlyFound;
                this.HierarchyStage = versePointerLink.IsChapter ? HierarchySearchManager.HierarchyStage.Page : HierarchySearchManager.HierarchyStage.ContentPlaceholder;
                this.HierarchyObjectInfo = new HierarchyObjectInfo()
                {                    
                    SectionId = versePointerLink.SectionId,
                    PageId = versePointerLink.PageId,
                    PageName = versePointerLink.PageName,
                    VerseInfo = new VerseObjectInfo(versePointerLink)                    
                };

                if (vp.IsMultiVerse &&
                                (findAllVerseObjects == FindVerseLevel.AllVerses || findAllVerseObjects == FindVerseLevel.OnlyVersesOfFirstChapter))
                {
                    List<VersePointer> verses;
                    if (findAllVerseObjects == FindVerseLevel.OnlyVersesOfFirstChapter)
                    {
                        Application temp = null;
                        verses = vp.GetAllIncludedVersesExceptFirst(ref temp,
                                                        new GetAllIncludedVersesExceptFirstArgs() { SearchOnlyForFirstChapter = true });
                    }
                    else
                        verses = vp.GetAllIncludedVersesExceptFirst(ref oneNoteApp,
                                                        new GetAllIncludedVersesExceptFirstArgs() { BibleNotebookId = bibleNotebookId });

                    foreach (var subVp in verses)
                    {
                        var subLink = OneNoteProxy.Instance.GetVersePointerLink(subVp);
                        if (subLink != null)
                            this.HierarchyObjectInfo.AdditionalObjectsIds.Add(subVp, new VerseObjectInfo(subLink));
                    }
                }
            }
        }       

        /// <summary>
        /// 
        /// </summary>
        /// <param name="oneNoteApp"></param>
        /// <param name="bibleNotebookId"></param>
        /// <param name="vp"></param>
        /// <param name="findAllVerseObjects">Поиск осуществляется только по текущей главе. Чтобы найти все стихи во всех главах (если например ссылка 3:4-6:8), то надо отдельно вызвать GetAllIncludedVersesExceptFirst</param>
        /// <returns></returns>
        public static HierarchySearchResult GetHierarchyObject(ref Application oneNoteApp, string bibleNotebookId, VersePointer vp, FindVerseLevel findAllVerseObjects)
        {   
            HierarchySearchResult result = new HierarchySearchResult();
            result.ResultType = HierarchySearchResultType.NotFound;

            if (!vp.IsValid)
                throw new ArgumentException("versePointer is not valid");

            if (OneNoteProxy.Instance.IsBibleVersesLinksCacheActive)
            {
                var simpleVersePointer = vp;
                var versePointerLink = OneNoteProxy.Instance.GetVersePointerLink(simpleVersePointer);
                if (versePointerLink != null)
                    return new HierarchySearchResult(ref oneNoteApp, bibleNotebookId, simpleVersePointer, versePointerLink, findAllVerseObjects);
            }

            XElement targetSection = FindBibleBookSection(ref oneNoteApp, bibleNotebookId, vp.Book.SectionName);
            if (targetSection != null)
            {
                result.HierarchyObjectInfo.SectionId = (string)targetSection.Attribute("ID");
                result.ResultType = HierarchySearchResultType.Successfully;
                result.HierarchyStage = HierarchyStage.Section;                

                XElement targetPage = HierarchySearchManager.FindPage(ref oneNoteApp, result.HierarchyObjectInfo.SectionId, vp.Chapter.Value);

                if (targetPage != null)
                {
                    result.HierarchyObjectInfo.PageId = (string)targetPage.Attribute("ID");
                    result.HierarchyObjectInfo.PageName = (string)targetPage.Attribute("name");
                    result.HierarchyStage = HierarchyStage.Page;

                    var pageContent = OneNoteProxy.Instance.GetPageContent(ref oneNoteApp, result.HierarchyObjectInfo.PageId, OneNoteProxy.PageType.Bible);
                    VerseNumber? verseNumber;
                    string verseTextWithoutNumber;
                    XElement targetVerseEl = FindVerse(pageContent.Content, vp.IsChapter, vp.Verse.Value, pageContent.Xnm, out verseNumber, out verseTextWithoutNumber);

                    if (targetVerseEl != null)
                    {
                        result.HierarchyObjectInfo.VerseInfo = new VerseObjectInfo()
                                                                        {
                                                                            ObjectId = (string)targetVerseEl.Parent.Attribute("objectID"),
                                                                            VerseNumber = verseNumber
                                                                        };
                        if (!vp.IsChapter)
                        {
                            result.HierarchyStage = HierarchyStage.ContentPlaceholder;

                            if (vp.IsMultiVerse &&
                                (findAllVerseObjects == FindVerseLevel.AllVerses || findAllVerseObjects == FindVerseLevel.OnlyVersesOfFirstChapter))
                            {
                                LoadAdditionalVersesInfo(ref oneNoteApp, bibleNotebookId, vp, findAllVerseObjects, pageContent, ref result);                              
                            }
                        }
                    }
                    else if (!vp.IsChapter)   // Если по идее должны были найти, а не нашли...
                    {
                        result.ResultType = HierarchySearchResultType.PartlyFound;
                    }                    
                }
            }

            return result;
        }

        private static void LoadAdditionalVersesInfo(ref Application oneNoteApp, string bibleNotebookId, VersePointer vp, FindVerseLevel findAllVerseObjects,
            OneNoteProxy.PageContent chapterPageContent, ref HierarchySearchResult result)
        {
            List<VersePointer> verses;
            if (findAllVerseObjects == FindVerseLevel.OnlyVersesOfFirstChapter)
            {
                Application temp = null;
                verses = vp.GetAllIncludedVersesExceptFirst(ref temp,
                                                new GetAllIncludedVersesExceptFirstArgs() { SearchOnlyForFirstChapter = true });
            }
            else
                verses = vp.GetAllIncludedVersesExceptFirst(ref oneNoteApp,
                                                new GetAllIncludedVersesExceptFirstArgs() { BibleNotebookId = bibleNotebookId });

            foreach (var additionalVerse in verses)
            {
                VerseNumber? additionalVerseNumber;
                string additionalVerseTextWithoutNumber;
                var additionalPageContent = chapterPageContent;
                if (additionalVerse.Chapter != vp.Chapter)
                {
                    var additionalPageEl = HierarchySearchManager.FindPage(ref oneNoteApp, result.HierarchyObjectInfo.SectionId, additionalVerse.Chapter.Value);
                    if (additionalPageEl == null)
                        continue;
                    var additionalPageId = (string)additionalPageEl.Attribute("ID");
                    additionalPageContent = OneNoteProxy.Instance.GetPageContent(ref oneNoteApp, additionalPageId, OneNoteProxy.PageType.Bible);
                }

                var additionalVeseEl = FindVerse(additionalPageContent.Content, vp.IsChapter, additionalVerse.Verse.Value, additionalPageContent.Xnm,
                    out additionalVerseNumber, out additionalVerseTextWithoutNumber);
                if (additionalVeseEl != null)
                {
                    result.HierarchyObjectInfo.AdditionalObjectsIds.Add(additionalVerse,
                                                                        new VerseObjectInfo()
                                                                        {
                                                                            ObjectId = (string)additionalVeseEl.Parent.Attribute("objectID"),
                                                                            VerseNumber = additionalVerseNumber
                                                                        });
                }
                else
                    continue;
            }
        }

        internal static XElement FindChapterPage(ref Application oneNoteApp, string bibleNotebookId, string sectionName, int chapterIndex)
        {
            XElement targetSection = FindBibleBookSection(ref oneNoteApp, bibleNotebookId, sectionName);
            if (targetSection != null)
            {
                return HierarchySearchManager.FindPage(ref oneNoteApp, (string)targetSection.Attribute("ID"), chapterIndex);                
            }

            return null;
        }        

        internal static XElement FindVerse(XDocument pageContent, bool isChapter, int verse, XmlNamespaceManager xnm, out VerseNumber? verseNumber, out string verseTextWithoutNumber)
        {
            XElement pointerElement = null;
            verseNumber = null;
            verseTextWithoutNumber = null;

            if (!isChapter)
            {
                string[] searchPatterns = new string[] { 
                    "/one:Page/one:Outline/one:OEChildren/one:OE/one:Table/one:Row/one:Cell[1]/one:OEChildren/one:OE/one:T[starts-with(.,'{0} ')]",
                    "/one:Page/one:Outline/one:OEChildren/one:OE/one:T[starts-with(.,'{0} ')]",
                    "/one:Page/one:Outline/one:OEChildren/one:OE/one:Table/one:Row/one:Cell[1]/one:OEChildren/one:OE/one:T[starts-with(.,'{0}&nbsp;')]",
                    "/one:Page/one:Outline/one:OEChildren/one:OE/one:T[starts-with(.,'{0}&nbsp;')]",
                    "/one:Page/one:Outline/one:OEChildren/one:OE/one:Table/one:Row/one:Cell[1]/one:OEChildren/one:OE/one:T[contains(.,'>{0}<')]",
                    "/one:Page/one:Outline/one:OEChildren/one:OE/one:T[contains(.,'>{0}<')]",
                    "/one:Page/one:Outline/one:OEChildren/one:OE/one:Table/one:Row/one:Cell[1]/one:OEChildren/one:OE/one:T[contains(.,'>{0}&nbsp;')]",
                    "/one:Page/one:Outline/one:OEChildren/one:OE/one:T[contains(.,'>{0}&nbsp;')]",
                    "/one:Page/one:Outline/one:OEChildren/one:OE/one:Table/one:Row/one:Cell[1]/one:OEChildren/one:OE/one:T[contains(.,'>{0} ')]",
                    "/one:Page/one:Outline/one:OEChildren/one:OE/one:T[contains(.,'>{0} ')]" };

                foreach (string pattern in searchPatterns)
                {
                    pointerElement = pageContent.XPathSelectElement(string.Format(pattern, verse), xnm);

                    if (pointerElement != null)
                    {                        
                        break;
                    }
                }

                verseNumber = null;
                if (pointerElement != null)                
                    verseNumber = VerseNumber.GetFromVerseText(pointerElement.Value, out verseTextWithoutNumber);                

                if (pointerElement == null || verseNumber == null || !verseNumber.Value.IsVerseBelongs(verse))
                    pointerElement = FindVerseWithIterate(pageContent, verse, xnm, out verseNumber, out verseTextWithoutNumber);                
            }
            else               // тогда возвращаем хотя бы ссылку на заголовок
            {
                pointerElement = NotebookGenerator.GetPageTitle(pageContent, xnm);
            }

            return pointerElement;
        }

        private static XElement FindVerseWithIterate(XDocument pageContent, int verse, XmlNamespaceManager xnm, 
            out VerseNumber? verseNumber, out string verseTextWithoutNumber)
        {
            verseNumber = null;
            verseTextWithoutNumber = null;

            foreach (var textEl in pageContent.Root.XPathSelectElements("one:Outline/one:OEChildren/one:OE/one:Table/one:Row/one:Cell[1]/one:OEChildren/one:OE/one:T", xnm)
                .Union(pageContent.Root.XPathSelectElements("one:Outline/one:OEChildren/one:OE/one:T", xnm)))
            {
                verseNumber = VerseNumber.GetFromVerseText(textEl.Value, out verseTextWithoutNumber);
                if (verseNumber.HasValue && verseNumber.Value.IsVerseBelongs(verse))
                    return textEl;
            }

            return null;
        }

        private static XElement FindPage(ref Application oneNoteApp, string sectionId, int chapter)
        {
            OneNoteProxy.HierarchyElement sectionDocument = OneNoteProxy.Instance.GetHierarchy(ref oneNoteApp, sectionId, HierarchyScope.hsPages);

            return FindChapterPage(sectionDocument.Content.Root, chapter, sectionDocument.Xnm);
        }

        internal static XElement FindChapterPage(XElement sectionPagesEl, int chapter, XmlNamespaceManager xnm)
        {
            XElement page = sectionPagesEl.XPathSelectElement(string.Format("one:Page[starts-with(@name,'{0} ') and @pageLevel=2]", chapter), xnm);  // "@pageLevel=2" - это условие нам надо, чтобы найти главу, например, "2 глава. 2 Петра", а не просто "2 Петра", 
            if (page == null)
                page = sectionPagesEl.XPathSelectElement(string.Format("one:Page[starts-with(@name,'{0} ')]", chapter), xnm);

            if (page == null)  // нужно для Псалтыря, потому что там главы называются, например "Псалом 5"
                page = sectionPagesEl.XPathSelectElement(
                    string.Format("one:Page[' {0}'=substring(@name,string-length(@name)-{1})]", chapter, chapter.ToString().Length), xnm);

            return page;
        }

        public static XElement FindBibleBookSection(ref Application oneNoteApp, string notebookId, string bookSectionName)
        {
            OneNoteProxy.HierarchyElement document = OneNoteProxy.Instance.GetHierarchy(ref oneNoteApp, notebookId, HierarchyScope.hsSections);

            XElement targetSection = document.Content.Root.XPathSelectElement(
                string.Format("{0}one:SectionGroup[{2}]/one:Section[@name='{1}']",
                    !string.IsNullOrEmpty(SettingsManager.Instance.SectionGroupId_Bible)
                        ? string.Format("one:SectionGroup[@ID='{0}']/", SettingsManager.Instance.SectionGroupId_Bible) 
                        : string.Empty,
                        bookSectionName, OneNoteUtils.NotInRecycleXPathCondition),
                document.Xnm);

            return targetSection;
        }

        public static int? GetChapterVersesCount(ref Application oneNoteApp, string bibleNotebookId, VersePointer versePointer)
        {
            int? result = null;

            if (bibleNotebookId == SettingsManager.Instance.NotebookId_Bible 
                && SettingsManager.Instance.CurrentModuleCached != null 
                && SettingsManager.Instance.CurrentModuleCached.Version >= Consts.Constants.ModulesWithXmlBibleMinVersion)
            {
                result = ModulesManager.GetChapterVersesCount(SettingsManager.Instance.CurrentBibleContentCached, versePointer);
            }

            if (result == null)
            {
                var chapterPageResult = GetHierarchyObject(ref oneNoteApp, bibleNotebookId, versePointer, FindVerseLevel.OnlyFirstVerse);
                if (chapterPageResult.ResultType != HierarchySearchResultType.NotFound 
                    && chapterPageResult.HierarchyStage == HierarchyStage.Page)
                {
                    var pageContent = OneNoteProxy.Instance.GetPageContent(ref oneNoteApp, chapterPageResult.HierarchyObjectInfo.PageId, OneNoteProxy.PageType.Bible);
                    var table = pageContent.Content.Root.XPathSelectElement("//one:Table", pageContent.Xnm);
                    if (table != null)
                    {
                        return table.XPathSelectElements("one:Row", pageContent.Xnm).Count();
                    }
                }
            }

            return result;
        }
    }
}
