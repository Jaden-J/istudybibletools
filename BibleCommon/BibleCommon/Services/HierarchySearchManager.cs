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

        public class HierarchyObjectInfo
        {            
            public string SectionId { get; set; }
            public string PageId { get; set; }
            public string ContentObjectId { get; set; }
            public List<string> AdditionalObjectsIds { get; set; }
            public List<string> GetAllObjectsIds()
            {
                var result = new List<string>() { ContentObjectId };
                result.AddRange(AdditionalObjectsIds);
                return result;
            }

            public HierarchyObjectInfo()
            {
                this.AdditionalObjectsIds = new List<string>();
            }
        }

        public enum HierarchySearchResultType
        {
            NotFound,            
            PartlyFound,  // например надо было найти стих, а нашли только страницу (если искали Быт 1:120)
            Successfully,            
        }
        
        public class HierarchySearchResult
        {
            public HierarchyObjectInfo HierarchyObjectInfo { get; set; } // дополнительная информация о найденном объекте            
            public HierarchyStage HierarchyStage { get; set; }
            public HierarchySearchResultType ResultType { get; set; }     

            public HierarchySearchResult()
            {
                HierarchyObjectInfo = new HierarchyObjectInfo();
            }
        }

        public static HierarchySearchResult GetHierarchyObject(Application oneNoteApp, string bibleNotebookId, VersePointer vp)
        {
            return GetHierarchyObject(oneNoteApp, bibleNotebookId, vp, false);
        }

        public static HierarchySearchResult GetHierarchyObject(Application oneNoteApp, string bibleNotebookId, VersePointer vp, bool findAllVerseObjects)
        {
            HierarchySearchResult result = new HierarchySearchResult();
            result.ResultType = HierarchySearchResultType.NotFound;

            if (!vp.IsValid)
                throw new ArgumentException("versePointer is not valid");

            XElement targetSection = FindBibleBookSection(oneNoteApp, bibleNotebookId, vp.Book.SectionName);
            if (targetSection != null)
            {
                result.HierarchyObjectInfo.SectionId = (string)targetSection.Attribute("ID");
                result.ResultType = HierarchySearchResultType.Successfully;
                result.HierarchyStage = HierarchyStage.Section;                

                XElement targetPage = HierarchySearchManager.FindPage(oneNoteApp, result.HierarchyObjectInfo.SectionId, vp.Chapter.Value);

                if (targetPage != null)
                {
                    result.HierarchyObjectInfo.PageId = (string)targetPage.Attribute("ID");
                    result.HierarchyStage = HierarchyStage.Page;

                    var pageContent = OneNoteProxy.Instance.GetPageContent(oneNoteApp, result.HierarchyObjectInfo.PageId, OneNoteProxy.PageType.Bible);
                    XElement targetVerseEl = FindVerse(pageContent.Content, vp.IsChapter, vp.Verse.Value, pageContent.Xnm);

                    if (targetVerseEl != null)
                    {
                        result.HierarchyObjectInfo.ContentObjectId = (string)targetVerseEl.Parent.Attribute("objectID");

                        if (!vp.IsChapter)
                        {
                            result.HierarchyStage = HierarchyStage.ContentPlaceholder;

                            if (findAllVerseObjects && vp.IsMultiVerse)
                            {
                                foreach (var additionalVerse in vp.GetAllIncludedVersesExceptFirst(null,
                                                                    new GetAllIncludedVersesExceptFirstArgs() { SearchOnlyForFirstChapter = true }))
                                {
                                    var additionalVeseEl = FindVerse(pageContent.Content, vp.IsChapter, additionalVerse.Verse.Value, pageContent.Xnm);
                                    if (additionalVeseEl != null)
                                        result.HierarchyObjectInfo.AdditionalObjectsIds.Add((string)additionalVeseEl.Parent.Attribute("objectID"));
                                    else
                                        break;
                                }
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

        internal static XElement FindChapterPage(Application oneNoteApp, string bibleNotebookId, string sectionName, int chapterIndex)
        {
            XElement targetSection = FindBibleBookSection(oneNoteApp, bibleNotebookId, sectionName);
            if (targetSection != null)
            {
                return HierarchySearchManager.FindPage(oneNoteApp, (string)targetSection.Attribute("ID"), chapterIndex);                
            }

            return null;
        }        

        internal static XElement FindVerse(XDocument pageContent, bool isChapter, int verse, XmlNamespaceManager xnm)
        {
            XElement pointerElement = null;

            if (!isChapter)
            {
                string[] searchPatterns = new string[] { 
                "/one:Page/one:Outline/one:OEChildren/one:OE/one:Table/one:Row/one:Cell/one:OEChildren/one:OE/one:T[starts-with(.,'{0} ')]",
                "/one:Page/one:Outline/one:OEChildren/one:OE/one:T[starts-with(.,'{0} ')]",
                "/one:Page/one:Outline/one:OEChildren/one:OE/one:Table/one:Row/one:Cell/one:OEChildren/one:OE/one:T[starts-with(.,'{0}&nbsp;')]",
                "/one:Page/one:Outline/one:OEChildren/one:OE/one:T[starts-with(.,'{0}&nbsp;')]",
                "/one:Page/one:Outline/one:OEChildren/one:OE/one:Table/one:Row/one:Cell/one:OEChildren/one:OE/one:T[contains(.,'>{0}<')]",
                "/one:Page/one:Outline/one:OEChildren/one:OE/one:T[contains(.,'>{0}<')]",
                "/one:Page/one:Outline/one:OEChildren/one:OE/one:Table/one:Row/one:Cell/one:OEChildren/one:OE/one:T[contains(.,'>{0}&nbsp;')]",
                "/one:Page/one:Outline/one:OEChildren/one:OE/one:T[contains(.,'>{0}&nbsp;')]",
                "/one:Page/one:Outline/one:OEChildren/one:OE/one:Table/one:Row/one:Cell/one:OEChildren/one:OE/one:T[contains(.,'>{0} ')]",
                "/one:Page/one:Outline/one:OEChildren/one:OE/one:T[contains(.,'>{0} ')]"
            };

                foreach (string pattern in searchPatterns)
                {
                    pointerElement = pageContent.XPathSelectElement(string.Format(pattern, verse), xnm);

                    if (pointerElement != null)
                    {                        
                        break;
                    }
                }
            }
            else               // тогда возвращаем хотя бы ссылку на заголовок
            {
                pointerElement = pageContent.Root.XPathSelectElement("one:Title/one:OE/one:T", xnm);
            }

            return pointerElement;
        }

        private static XElement FindPage(Application oneNoteApp, string sectionId, int chapter)
        {
            OneNoteProxy.HierarchyElement sectionDocument = OneNoteProxy.Instance.GetHierarchy(oneNoteApp, sectionId, HierarchyScope.hsPages);

            return FindChapterPage(oneNoteApp, sectionDocument.Content.Root, chapter, sectionDocument.Xnm);
        }

        internal static XElement FindChapterPage(Application oneNoteApp, XElement sectionPagesEl, int chapter, XmlNamespaceManager xnm)
        {
            XElement page = sectionPagesEl.XPathSelectElement(string.Format("one:Page[starts-with(@name,'{0} ') and @pageLevel=2]", chapter), xnm);  // "@pageLevel=2" - это условие нам надо, чтобы найти главу, например, "2 глава. 2 Петра", а не просто "2 Петра", 
            if (page == null)
                page = sectionPagesEl.XPathSelectElement(string.Format("one:Page[starts-with(@name,'{0} ')]", chapter), xnm);

            if (page == null)  // нужно для Псалтыря, потому что там главы называются, например "Псалом 5"
                page = sectionPagesEl.XPathSelectElement(
                    string.Format("one:Page[' {0}'=substring(@name,string-length(@name)-{1})]", chapter, chapter.ToString().Length), xnm);

            return page;
        }

        public static XElement FindBibleBookSection(Application oneNoteApp, string notebookId, string bookSectionName)
        {
            OneNoteProxy.HierarchyElement document = OneNoteProxy.Instance.GetHierarchy(oneNoteApp, notebookId, HierarchyScope.hsSections);

            XElement targetSection = document.Content.Root.XPathSelectElement(
                string.Format("{0}one:SectionGroup[{2}]/one:Section[@name='{1}']",
                    !string.IsNullOrEmpty(SettingsManager.Instance.SectionGroupId_Bible)
                        ? string.Format("one:SectionGroup[@ID='{0}']/", SettingsManager.Instance.SectionGroupId_Bible) 
                        : string.Empty,
                        bookSectionName, OneNoteUtils.NotInRecycleXPathCondition),
                document.Xnm);

            return targetSection;
        }

        public static int? GetChapterVersesCount(Application oneNoteApp, string bibleNotebookId, VersePointer versePointer)
        {
            int? result = null;

            var chapterPageResult = GetHierarchyObject(oneNoteApp, bibleNotebookId, versePointer);
            if (chapterPageResult.ResultType != HierarchySearchResultType.NotFound)
            {
                var pageContent = OneNoteProxy.Instance.GetPageContent(oneNoteApp, chapterPageResult.HierarchyObjectInfo.PageId, OneNoteProxy.PageType.Bible);
                var table = pageContent.Content.Root.XPathSelectElement("//one:Table", pageContent.Xnm);
                if (table != null)
                {
                    return table.XPathSelectElements("one:Row", pageContent.Xnm).Count();
                }
            }

            return result;
        }
    }
}
