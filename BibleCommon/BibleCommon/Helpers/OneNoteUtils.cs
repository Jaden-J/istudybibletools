using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Xml;
using Microsoft.Office.Interop.OneNote;
using System.Reflection;
using System.IO;
using BibleCommon.Consts;
using BibleCommon.Services;
using System.Xml.XPath;
using System.Runtime.InteropServices;

namespace BibleCommon.Helpers
{
    public static class OneNoteUtils
    {
        public static bool NotebookExists(Application oneNoteApp, string notebookId)
        {
            OneNoteProxy.HierarchyElement hierarchy = OneNoteProxy.Instance.GetHierarchy(oneNoteApp, null, HierarchyScope.hsNotebooks);
            XElement bibleNotebook = hierarchy.Content.Root.XPathSelectElement(string.Format("one:Notebook[@ID='{0}']", notebookId), hierarchy.Xnm);
            return bibleNotebook != null;            
        }

        public static bool RootSectionGroupExists(Application oneNoteApp, string notebookId, string sectionGroupId)
        {
            OneNoteProxy.HierarchyElement hierarchy = OneNoteProxy.Instance.GetHierarchy(oneNoteApp, notebookId, HierarchyScope.hsChildren);
            XElement sectionGroup = hierarchy.Content.Root.XPathSelectElement(string.Format("one:SectionGroup[@ID='{0}']", sectionGroupId), hierarchy.Xnm);
            return sectionGroup != null;
        }

        public static string GetNotebookIdByName(Application oneNoteApp, string notebookName, bool refreshCache)
        {
            OneNoteProxy.HierarchyElement hierarchy = OneNoteProxy.Instance.GetHierarchy(oneNoteApp, null, HierarchyScope.hsNotebooks, refreshCache);
            XElement bibleNotebook = hierarchy.Content.Root.XPathSelectElement(string.Format("one:Notebook[@nickname='{0}']", notebookName), hierarchy.Xnm);
            if (bibleNotebook == null)
                bibleNotebook = hierarchy.Content.Root.XPathSelectElement(string.Format("one:Notebook[@name='{0}']", notebookName), hierarchy.Xnm);
            if (bibleNotebook != null)
            {
                return (string)bibleNotebook.Attribute("ID");
            }

            return string.Empty;
        }     

        public static string GetHierarchyElementName(Application oneNoteApp, string elementId)
        {   
            OneNoteProxy.HierarchyElement doc = OneNoteProxy.Instance.GetHierarchy(oneNoteApp, elementId, HierarchyScope.hsSelf);
            return (string)doc.Content.Root.Attribute("name");
        }

        public static XDocument GetXDocument(string xml, out XmlNamespaceManager xnm)
        {
            XDocument xd = XDocument.Parse(xml);
            xnm = new XmlNamespaceManager(new NameTable());
            xnm.AddNamespace("one", Constants.OneNoteXmlNs);
            return xd;
        }

        // возвращает количество родительских узлов
        public static int GetDepthLevel(XElement element)
        {
            int result = 0;

            if (element.Parent != null)
            {                
                result += 1 + GetDepthLevel(element.Parent);
            }

            return result;
        }

        public static bool IsRecycleBin(XElement hierarchyElement)
        {
            return bool.Parse(GetAttributeValue(hierarchyElement, "isInRecycleBin", false.ToString()))
                || bool.Parse(GetAttributeValue(hierarchyElement, "isRecycleBin", false.ToString()));
        }

        public static string NotInRecycleXPathCondition
        {
            get
            {
                return "not(@isInRecycleBin) and not(@isRecycleBin)";
            }
        }

        public static string GetAttributeValue(XElement el, string attributeName, string defaultValue)
        {
            if (el.Attribute(attributeName) != null)
            {
                return (string)el.Attribute(attributeName).Value;                
            }

            return defaultValue;
        }


        public static string GenerateHref(Application oneNoteApp, string title, string pageId, string objectId)        
        {
            string link = OneNoteProxy.Instance.GenerateHref(oneNoteApp, pageId, objectId);            

            return string.Format("<a href=\"{0}\">{1}</a>", link, title);
        }

        public static void NormalizaTextElement(XElement textElement)  // must be one:T element
        {
            if (textElement != null)
            {
                if (!string.IsNullOrEmpty(textElement.Value))
                {
                    textElement.Value = textElement.Value.Replace("\n", " ");
                }
            }
        }

        public static void UpdatePageContentSafe(Application oneNoteApp, XDocument pageContent)
        {
            try
            {
                oneNoteApp.UpdatePageContent(pageContent.ToString());
            }
            catch (COMException ex)
            {
                if (ex.ErrorCode == -2147213304)
                    throw new Exception(Constants.Error_UpdateError_InksOnPages);
                else
                    throw;
            }
        }

        public static void UpdatePageMetaData(Application oneNoteApp, XElement pageContent, string key, string value, XmlNamespaceManager xnm)
        {
            var metaElement = pageContent.XPathSelectElement(string.Format("one:Meta[@name='{0}']", key), xnm);
            if (metaElement != null)
            {
                metaElement.SetAttributeValue("content", value);
            }
            else
            {
                XNamespace nms = XNamespace.Get(Constants.OneNoteXmlNs);

                var pageSettings = pageContent.XPathSelectElement("one:PageSettings", xnm);

                var meta = new XElement(nms + "Meta",
                                            new XAttribute("name", key),
                                            new XAttribute("content", value));


                if (pageSettings != null)
                    pageSettings.AddBeforeSelf(meta);
                else
                    pageContent.AddFirst(meta);
            }
        }


        public static string GetPageMetaData(Application oneNoteApp, XElement pageContent, string key, XmlNamespaceManager xnm)
        {
            var metaElement = pageContent.XPathSelectElement(string.Format("one:Meta[@name='{0}']", key), xnm);
            if (metaElement != null)
            {
                return metaElement.Attribute("content").Value;
            }

            return string.Empty;
        }

        public static  NotebookIterator.PageInfo GetCurrentPageInfo(Application oneNoteApp)
        {
            if (oneNoteApp.Windows.CurrentWindow == null)
                throw new Exception("Не найдено открытой записной книжки");

            string currentPageId = oneNoteApp.Windows.CurrentWindow.CurrentPageId;
            if (string.IsNullOrEmpty(currentPageId))
                throw new Exception("Не найдено открытой страницы заметок");

            string currentSectionId = oneNoteApp.Windows.CurrentWindow.CurrentSectionId;
            string currentSectionGroupId = oneNoteApp.Windows.CurrentWindow.CurrentSectionGroupId;
            string currentNotebookId = oneNoteApp.Windows.CurrentWindow.CurrentNotebookId;


            return new NotebookIterator.PageInfo()
            {
                SectionGroupId = currentSectionGroupId,
                SectionId = currentSectionId,
                Id = currentPageId
            };
        }
    }
}
