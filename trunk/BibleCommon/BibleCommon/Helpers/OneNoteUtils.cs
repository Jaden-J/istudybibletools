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
using System.Threading;
using BibleCommon.Common;

namespace BibleCommon.Helpers
{
    public static class OneNoteUtils
    {
        private static bool? _isOneNote2010;
        public static bool IsOneNote2010Cached(Application oneNoteApp)
        {
            if (!_isOneNote2010.HasValue)
            {
                if (oneNoteApp != null)
                    _isOneNote2010 = oneNoteApp.GetType().Assembly.GetName().Version < new Version(15, 0, 0, 0);
                else
                    return true;
            }

            return _isOneNote2010.Value;
        }

        public static bool NotebookExists(Application oneNoteApp, string notebookId, bool refreshCache = false)
        {
            OneNoteProxy.HierarchyElement hierarchy = OneNoteProxy.Instance.GetHierarchy(oneNoteApp, null, HierarchyScope.hsNotebooks, refreshCache);
            XElement bibleNotebook = hierarchy.Content.Root.XPathSelectElement(string.Format("one:Notebook[@ID='{0}']", notebookId), hierarchy.Xnm);
            return bibleNotebook != null;            
        }

        public static void CloseNotebookSafe(Application oneNoteApp, string notebookId)
        {
            if (NotebookExists(oneNoteApp, notebookId, true))
                oneNoteApp.CloseNotebook(notebookId);
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

        public static string GetHierarchyElementNickname(Application oneNoteApp, string elementId)
        {
            OneNoteProxy.HierarchyElement doc = OneNoteProxy.Instance.GetHierarchy(oneNoteApp, elementId, HierarchyScope.hsSelf);
            return (string)doc.Content.Root.Attribute("nickname");
        }

        public static string GetHierarchyElementName(Application oneNoteApp, string elementId)
        {   
            OneNoteProxy.HierarchyElement doc = OneNoteProxy.Instance.GetHierarchy(oneNoteApp, elementId, HierarchyScope.hsSelf);
            return (string)doc.Content.Root.Attribute("name");
        }

        public static XmlNamespaceManager GetOneNoteXNM()
        {
            var xnm = new XmlNamespaceManager(new NameTable());
            xnm.AddNamespace("one", Constants.OneNoteXmlNs);

            return xnm;
        }

        public static XDocument GetXDocument(string xml, out XmlNamespaceManager xnm)
        {
            XDocument xd = XDocument.Parse(xml);
            xnm = GetOneNoteXNM();
            return xd;
        }

        public static bool HierarchyElementExists(Application oneNoteApp, string hierarchyId)
        {
            try
            {
                string xml;
                oneNoteApp.GetHierarchy(hierarchyId, HierarchyScope.hsSelf, out xml, Constants.CurrentOneNoteSchema);
                return true;
            }
            catch (COMException ex)
            {
                if (ex.Message.Contains(Utils.GetHexError(Error.hrObjectDoesNotExist)))
                    return false;
                else
                    throw;
            }
        }

        public static XDocument GetHierarchyElement(Application oneNoteApp, string hierarchyId, HierarchyScope scope, out XmlNamespaceManager xnm)
        {
            string xml;
            oneNoteApp.GetHierarchy(hierarchyId, scope, out xml, Constants.CurrentOneNoteSchema);
            return GetXDocument(xml, out xnm);            
        }


        public static XDocument GetPageContent(Application oneNoteApp, string pageId, out XmlNamespaceManager xnm)
        {
            return GetPageContent(oneNoteApp, pageId, PageInfo.piBasic, out xnm);
        }

        public static XDocument GetPageContent(Application oneNoteApp, string pageId, PageInfo pageInfo, out XmlNamespaceManager xnm)
        {
            string xml;
            oneNoteApp.GetPageContent(pageId, out xml, pageInfo, Constants.CurrentOneNoteSchema);            
            return OneNoteUtils.GetXDocument(xml, out xnm);
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
                return (string)el.Attribute(attributeName);                
            }

            return defaultValue;
        }


        public static string GetOrGenerateHref(Application oneNoteApp, string title, string objectHref, string pageId, string objectId, params string[] additionalLinkQueryParameters)
        {
            string link;

            if (!string.IsNullOrEmpty(objectHref))
                link = objectHref;
            else
                link = OneNoteProxy.Instance.GenerateHref(oneNoteApp, pageId, objectId);

            foreach (var param in additionalLinkQueryParameters)
                link += "&" + param;

            return string.Format("<a href=\"{0}\">{1}</a>", link, title);            
        }

        public static string GenerateHref(Application oneNoteApp, string title, string pageId, string objectId)        
        {
            return GetOrGenerateHref(oneNoteApp, title, null, pageId, objectId);
        }

        public static XElement NormalizeTextElement(XElement textElement)  // must be one:T element
        {
            if (textElement != null)
            {
                if (!string.IsNullOrEmpty(textElement.Value))
                {
                    textElement.Value = textElement.Value.Replace("\n", " ").Replace("&nbsp;", " ");
                }
            }

            return textElement;
        }


        public static void UpdatePageContentSafe(ref Application oneNoteApp, XDocument pageContent, XmlNamespaceManager xnm)
        {
            UpdatePageContentSafeInternal(ref oneNoteApp, pageContent, xnm, 0);
        }

        private static void UpdatePageContentSafeInternal(ref Application oneNoteApp, XDocument pageContent, XmlNamespaceManager xnm, int attemptCount)
        {
            var inkNodes = pageContent.Root.XPathSelectElements("one:InkDrawing", xnm)
                            //.Union(doc.Root.XPathSelectElements("//one:OE[.//one:InkDrawing]", xnm))    // тогда удалятся все неподдерживаемые элементы. Но тогда у пользователей будут просто удаляться некоторые рисунки
                            .Union(pageContent.Root.XPathSelectElements("one:Outline[.//one:InkWord]", xnm)).ToArray();
            foreach (var inkNode in inkNodes)
                inkNode.Remove();

            try
            {
                oneNoteApp.UpdatePageContent(pageContent.ToString(), DateTime.MinValue, Constants.CurrentOneNoteSchema);                
            }
            catch (COMException ex)
            {
                if (ex.ErrorCode == -2147213304)
                    throw new Exception(Resources.Constants.Error_UpdateError_InksOnPages);

                if (ex.Message.Contains("0x80010100") || ex.Message.Contains("0x800706BA") || ex.Message.Contains("0x800706BE"))  // "System.Runtime.InteropServices.COMException (0x80010100): System call failed. (Exception from HRESULT: 0x80010100 (RPC_E_SYS_CALL_FAILED))"
                {
                    Logger.LogMessageSilient("UpdatePageContentSafeInternal. Attempt {0}: {1}", attemptCount, ex.Message);
                    if (attemptCount <= 10)
                    {
                        attemptCount++;
                        Thread.Sleep(1000 * attemptCount);
                        oneNoteApp = null;
                        oneNoteApp = new Application();
                        UpdatePageContentSafeInternal(ref oneNoteApp, pageContent, xnm, attemptCount);
                    }
                    else
                        throw;
                }
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

                var pageSettings = pageContent.XPathSelectElement("one:MediaPlaylist", xnm) ?? pageContent.XPathSelectElement("one:PageSettings", xnm);
                
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
                return (string)metaElement.Attribute("content");
            }

            return null;
        }    

        public static NotebookIterator.PageInfo GetCurrentPageInfo(Application oneNoteApp)
        {
            if (oneNoteApp.Windows.CurrentWindow == null)
                throw new ProgramException(BibleCommon.Resources.Constants.Error_OpenedNotebookNotFound);

            string currentPageId = oneNoteApp.Windows.CurrentWindow.CurrentPageId;
            if (string.IsNullOrEmpty(currentPageId))
                throw new ProgramException(BibleCommon.Resources.Constants.Error_OpenedNotePageNotFound);

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

        public static Dictionary<string, string> GetExistingNotebooks(Application oneNoteApp)
        {
            Dictionary<string, string> result = new Dictionary<string, string>();

            OneNoteProxy.HierarchyElement hierarchy = OneNoteProxy.Instance.GetHierarchy(oneNoteApp, null, HierarchyScope.hsNotebooks, true);

            foreach (XElement notebook in hierarchy.Content.Root.XPathSelectElements("one:Notebook", hierarchy.Xnm))
            {
                string name = (string)notebook.Attribute("nickname");
                if (string.IsNullOrEmpty(name))
                    name = (string)notebook.Attribute("name");
                string id = (string)notebook.Attribute("ID");
                result.Add(id, name);
            }

            return result;
        }
    }
}
