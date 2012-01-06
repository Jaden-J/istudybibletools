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

        public static HierarchySearchResult GetHierarchyObject(Application oneNoteApp, string notebookId, VersePointer vp)
        {
            HierarchySearchResult result = new HierarchySearchResult();
            result.ResultType = HierarchySearchResultType.NotFound;

            XElement targetSection = FindSection(oneNoteApp, notebookId, vp);
            if (targetSection != null)
            {
                result.HierarchyObjectInfo.SectionId = (string)targetSection.Attribute("ID");
                result.ResultType = HierarchySearchResultType.Successfully;
                result.HierarchyStage = HierarchyStage.Section;                

                XElement targetPage = HierarchySearchManager.FindPage(oneNoteApp, result.HierarchyObjectInfo.SectionId, vp);

                if (targetPage != null)
                {
                    result.HierarchyObjectInfo.PageId = (string)targetPage.Attribute("ID");
                    result.HierarchyStage = HierarchyStage.Page;

                    XElement targetVerse = HierarchySearchManager.FindVerse(oneNoteApp, result.HierarchyObjectInfo.PageId, vp);
                    if (targetVerse != null)
                    {
                        result.HierarchyObjectInfo.ContentObjectId = (string)targetVerse.Parent.Attribute("objectID");

                        if (!vp.IsChapter)
                            result.HierarchyStage = HierarchyStage.ContentPlaceholder;
                    }
                    else if (!vp.IsChapter)   // Если по идее должны были найти, а не нашли...
                    {
                        result.ResultType = HierarchySearchResultType.PartlyFound;
                    }                    
                }
            }

            return result;
        }

        private static XElement FindVerse(Application oneNoteApp, string pageId, VersePointer vp)
        {
            string pageContentXml = OneNoteProxy.Instance.GetPageContent(oneNoteApp, pageId);            

            XmlNamespaceManager xnm;
            XDocument document = Utils.GetXDocument(pageContentXml, out xnm);

            XElement pointerElement = null;

            if (!vp.IsChapter)
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
                    pointerElement = document.XPathSelectElement(string.Format(pattern, vp.Verse), xnm);

                    if (pointerElement != null)
                    {                        
                        break;
                    }
                }
            }
            else               // тогда возвращаем хотя бы ссылку на заголовок
            {
                pointerElement = document.Root.XPathSelectElement("one:Title/one:OE/one:T", xnm);
            }

            return pointerElement;
        }


        private static XElement FindPage(Application oneNoteApp, string sectionId, VersePointer vp)
        {
            string sectionContentXml = OneNoteProxy.Instance.GetHierarchy(oneNoteApp, sectionId, HierarchyScope.hsPages);

            XmlNamespaceManager xnm;
            XDocument sectionDocument = Utils.GetXDocument(sectionContentXml, out xnm);

            XElement page = sectionDocument.Root.XPathSelectElement(string.Format("one:Page[starts-with(@name,'{0} ')]", vp.Chapter), xnm);

            if (page == null)  // нужно для Псалтыря, потому что там главы называются, например "Псалом 5"
                page = sectionDocument.Root.XPathSelectElement(string.Format("one:Page[' {0}'=substring(@name,string-length(@name)-{1})]", vp.Chapter, vp.Chapter.ToString().Length), xnm);

            return page;
        }

        private static XElement FindSection(Application oneNoteApp, string notebookId, VersePointer vp)
        {
            string notebookContentXml = OneNoteProxy.Instance.GetHierarchy(oneNoteApp, notebookId, HierarchyScope.hsSections);

            XmlNamespaceManager xnm;
            XDocument document = Utils.GetXDocument(notebookContentXml, out xnm);

            XElement targetSection = document.Root.XPathSelectElement(
            string.Format("one:SectionGroup[@name='{0}']/one:SectionGroup/one:Section[@name='{1}']",
            SettingsManager.Instance.BibleSectionGroupName, vp.BookName), xnm);

            return targetSection;
        }      
    }
}
