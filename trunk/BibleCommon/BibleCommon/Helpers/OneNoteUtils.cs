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

        public static bool NotebookExists(ref Application oneNoteApp, string notebookId, bool refreshCache = false)
        {
            OneNoteProxy.HierarchyElement hierarchy = OneNoteProxy.Instance.GetHierarchy(ref oneNoteApp, null, HierarchyScope.hsNotebooks, refreshCache);
            XElement bibleNotebook = hierarchy.Content.Root.XPathSelectElement(string.Format("one:Notebook[@ID='{0}']", notebookId), hierarchy.Xnm);
            return bibleNotebook != null;            
        }

        public static void CloseNotebookSafe(ref Application oneNoteApp, string notebookId)
        {
            if (NotebookExists(ref oneNoteApp, notebookId, true))
            {
                OneNoteUtils.UseOneNoteAPI(ref oneNoteApp, (oneNoteAppSafe) =>
                {
                    oneNoteAppSafe.CloseNotebook(notebookId);
                });
            }
        }

        public static bool RootSectionGroupExists(ref Application oneNoteApp, string notebookId, string sectionGroupId)
        {
            OneNoteProxy.HierarchyElement hierarchy = OneNoteProxy.Instance.GetHierarchy(ref oneNoteApp, notebookId, HierarchyScope.hsChildren);
            XElement sectionGroup = hierarchy.Content.Root.XPathSelectElement(string.Format("one:SectionGroup[@ID='{0}']", sectionGroupId), hierarchy.Xnm);
            return sectionGroup != null;
        }

        public static string GetNotebookIdByName(ref Application oneNoteApp, string notebookName, bool refreshCache)
        {
            OneNoteProxy.HierarchyElement hierarchy = OneNoteProxy.Instance.GetHierarchy(ref oneNoteApp, null, HierarchyScope.hsNotebooks, refreshCache);
            XElement bibleNotebook = hierarchy.Content.Root.XPathSelectElement(string.Format("one:Notebook[@nickname='{0}']", notebookName), hierarchy.Xnm);
            if (bibleNotebook == null)
                bibleNotebook = hierarchy.Content.Root.XPathSelectElement(string.Format("one:Notebook[@name='{0}']", notebookName), hierarchy.Xnm);
            if (bibleNotebook != null)
            {
                return (string)bibleNotebook.Attribute("ID");
            }

            return string.Empty;
        }

        public static string GetHierarchyElementNickname(ref Application oneNoteApp, string elementId)
        {
            OneNoteProxy.HierarchyElement doc = OneNoteProxy.Instance.GetHierarchy(ref oneNoteApp, elementId, HierarchyScope.hsSelf);
            return (string)doc.Content.Root.Attribute("nickname");
        }

        public static string GetHierarchyElementName(ref Application oneNoteApp, string elementId)
        {   
            OneNoteProxy.HierarchyElement doc = OneNoteProxy.Instance.GetHierarchy(ref oneNoteApp, elementId, HierarchyScope.hsSelf);
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

        public static bool HierarchyElementExists(ref Application oneNoteApp, string hierarchyId)
        {
            try
            {
                string xml = null;

                UseOneNoteAPI(ref oneNoteApp, (oneNoteAppSafe) =>
                {
                    oneNoteAppSafe.GetHierarchy(hierarchyId, HierarchyScope.hsSelf, out xml, Constants.CurrentOneNoteSchema);
                });

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

        public static XDocument GetHierarchyElement(ref Application oneNoteApp, string hierarchyId, HierarchyScope scope, out XmlNamespaceManager xnm)
        {
            string xml = null;
            UseOneNoteAPI(ref oneNoteApp, (oneNoteAppSafe) =>
            {
                oneNoteAppSafe.GetHierarchy(hierarchyId, scope, out xml, Constants.CurrentOneNoteSchema);
            });
            return GetXDocument(xml, out xnm);            
        }

        public static XElement GetHierarchyElementByName(ref Application oneNoteApp, string elementTag, string elementName, string parentElementId)
        {
            XmlNamespaceManager xnm;
            var parentEl = GetHierarchyElement(ref oneNoteApp, parentElementId, HierarchyScope.hsChildren, out xnm);

            return parentEl.Root.XPathSelectElement(string.Format("one:{0}[@name='{1}']", elementTag, elementName), xnm);
        }


        public static XDocument GetPageContent(ref Application oneNoteApp, string pageId, out XmlNamespaceManager xnm)
        {
            return GetPageContent(ref oneNoteApp, pageId, PageInfo.piBasic, out xnm);
        }

        public static XDocument GetPageContent(ref Application oneNoteApp, string pageId, PageInfo pageInfo, out XmlNamespaceManager xnm)
        {
            string xml = null;

            UseOneNoteAPI(ref oneNoteApp, (oneNoteAppSafe) =>
            {
                oneNoteAppSafe.GetPageContent(pageId, out xml, pageInfo, Constants.CurrentOneNoteSchema);
            });

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


        public static string GetOrGenerateHref(ref Application oneNoteApp, string title, string objectHref, string pageId, string objectId, params string[] additionalLinkQueryParameters)
        {
            string link;

            if (!string.IsNullOrEmpty(objectHref))
                link = objectHref;
            else
                link = OneNoteProxy.Instance.GenerateHref(ref oneNoteApp, pageId, objectId);

            foreach (var param in additionalLinkQueryParameters)
                link += "&" + param;

            return string.Format("<a href=\"{0}\">{1}</a>", link, title);            
        }

        public static string GenerateHref(ref Application oneNoteApp, string title, string pageId, string objectId)        
        {
            return GetOrGenerateHref(ref oneNoteApp, title, null, pageId, objectId);
        }

        public static XElement NormalizeTextElement(XElement textElement)  // must be one:T element
        {
            if (textElement != null)
            {
                if (!string.IsNullOrEmpty(textElement.Value))
                {
                    textElement.Value = textElement.Value.Replace("\n", " ").Replace("&nbsp;", " ").Replace("<br>", "<br>\n");
                }
            }

            return textElement;
        }


        public static void UpdatePageContentSafe(ref Application oneNoteApp, XDocument pageContent, XmlNamespaceManager xnm)
        {
            UpdatePageContentSafeInternal(ref oneNoteApp, pageContent, xnm, 0);
        }

        private static void UpdatePageContentSafeInternal(ref Application oneNoteApp, XDocument pageContent, XmlNamespaceManager xnm, int attemptsCount)
        {
            var inkNodes = pageContent.Root.XPathSelectElements("one:InkDrawing", xnm)
                            .Union(pageContent.Root.XPathSelectElements("//one:OE[.//one:InkDrawing]", xnm))    
                            .Union(pageContent.Root.XPathSelectElements("one:Outline[.//one:InkWord]", xnm)).ToArray();

            foreach (var inkNode in inkNodes)
            {
                if (inkNode.XPathSelectElement(".//one:T", xnm) == null)
                    inkNode.Remove();
                else
                {
                    var inkWords = inkNode.XPathSelectElements(".//one:InkWord", xnm).Where(ink => ink.XPathSelectElement(".//one:CallbackID", xnm) == null).ToArray();
                    inkWords.Remove();                 
                }
            }

            try
            {
                UseOneNoteAPI(ref oneNoteApp, (oneNoteAppSafe) =>
                {
                    oneNoteAppSafe.UpdatePageContent(pageContent.ToString(), DateTime.MinValue, Constants.CurrentOneNoteSchema);
                });
            }
            catch (COMException ex)
            {
                if (ex.Message.Contains(Utils.GetHexError(Error.hrInsertingInk)))
                    throw new Exception(Resources.Constants.Error_UpdateError_InksOnPages);               
                else
                    throw;
            }
        }

        public static void UseOneNoteAPI(ref Application oneNoteApp, Action action)
        {
            UseOneNoteAPI(ref oneNoteApp, (safeOneNoteApp) => { action(); });
        }

        public static void UseOneNoteAPI(ref Application oneNoteApp, Action<IApplication> action)
        {
            UseOneNoteAPIInternal(ref oneNoteApp, action, 0);
        }

        private static void UseOneNoteAPIInternal(ref Application oneNoteApp, Action<IApplication> action, int attemptsCount)
        {
            try
            {
                action(oneNoteApp);
            }
            catch (COMException ex)
            {
                if (ex.Message.Contains("0x80010100") || ex.Message.Contains("0x800706BA") || ex.Message.Contains("0x800706BE"))  // "System.Runtime.InteropServices.COMException (0x80010100): System call failed. (Exception from HRESULT: 0x80010100 (RPC_E_SYS_CALL_FAILED))"
                {
                    Logger.LogMessageSilientParams("UseOneNoteAPI. Attempt {0}: {1}", attemptsCount, ex.Message);
                    if (attemptsCount <= 15)
                    {
                        attemptsCount++;
                        Thread.Sleep(1000 * attemptsCount);
                        System.Windows.Forms.Application.DoEvents();
                        oneNoteApp = null;
                        oneNoteApp = new Application();
                        UseOneNoteAPIInternal(ref oneNoteApp, action, attemptsCount);
                    }
                    else
                        throw;
                }
                else
                    throw;
            }
        }

        public static void UpdatePageMetaData(XElement pageContent, string key, string value, XmlNamespaceManager xnm)
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


        public static string GetPageMetaData(XElement pageContent, string key, XmlNamespaceManager xnm)
        {
            var metaElement = pageContent.XPathSelectElement(string.Format("one:Meta[@name='{0}']", key), xnm);
            if (metaElement != null)
            {
                return (string)metaElement.Attribute("content");
            }

            return null;
        }    

        public static NotebookIterator.PageInfo GetCurrentPageInfo(ref Application oneNoteApp)
        {
            string currentPageId = null;
            string currentSectionId = null;
            string currentSectionGroupId = null;
            string currentNotebookId = null;

            UseOneNoteAPI(ref oneNoteApp, (oneNoteAppSafe) =>
            {
                if (oneNoteAppSafe.Windows.CurrentWindow == null)
                    throw new ProgramException(BibleCommon.Resources.Constants.Error_OpenedNotebookNotFound);

                currentPageId = oneNoteAppSafe.Windows.CurrentWindow.CurrentPageId;
                if (string.IsNullOrEmpty(currentPageId))
                    throw new ProgramException(BibleCommon.Resources.Constants.Error_OpenedNotePageNotFound);

                currentSectionId = oneNoteAppSafe.Windows.CurrentWindow.CurrentSectionId;
                currentSectionGroupId = oneNoteAppSafe.Windows.CurrentWindow.CurrentSectionGroupId;
                currentNotebookId = oneNoteAppSafe.Windows.CurrentWindow.CurrentNotebookId;
            });

            return new NotebookIterator.PageInfo()
            {
                SectionGroupId = currentSectionGroupId,
                SectionId = currentSectionId,
                Id = currentPageId
            };
        }

        public static Dictionary<string, string> GetExistingNotebooks(ref Application oneNoteApp)
        {
            Dictionary<string, string> result = new Dictionary<string, string>();

            OneNoteProxy.HierarchyElement hierarchy = OneNoteProxy.Instance.GetHierarchy(ref oneNoteApp, null, HierarchyScope.hsNotebooks, true);

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
