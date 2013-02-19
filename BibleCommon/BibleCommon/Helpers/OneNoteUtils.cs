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
using System.Text.RegularExpressions;
using System.Globalization;

namespace BibleCommon.Helpers
{
    public static class OneNoteUtils
    {
        private static bool? _isOneNote2010;
        public static bool IsOneNote2010Cached
        {
            get
            {
                return false;
                if (!_isOneNote2010.HasValue)
                {
                    var assembly = new Application().GetType().Assembly;
                    _isOneNote2010 = assembly.GetName().Version < new Version(15, 0, 0, 0);
                }

                return _isOneNote2010.Value;
            }
        }

        public static bool NotebookExists(ref Application oneNoteApp, string notebookId, bool refreshCache = false)
        {
            OneNoteProxy.HierarchyElement hierarchy = OneNoteProxy.Instance.GetHierarchy(ref oneNoteApp, null, HierarchyScope.hsNotebooks, refreshCache);
            XElement bibleNotebook = hierarchy.Content.Root.XPathSelectElement(string.Format("one:Notebook[@ID=\"{0}\"]", notebookId), hierarchy.Xnm);
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
            XElement sectionGroup = hierarchy.Content.Root.XPathSelectElement(string.Format("one:SectionGroup[@ID=\"{0}\"]", sectionGroupId), hierarchy.Xnm);
            return sectionGroup != null;
        }

        public static string GetNotebookIdByName(ref Application oneNoteApp, string notebookName, bool refreshCache)
        {
            OneNoteProxy.HierarchyElement hierarchy = OneNoteProxy.Instance.GetHierarchy(ref oneNoteApp, null, HierarchyScope.hsNotebooks, refreshCache);
            XElement bibleNotebook = hierarchy.Content.Root.XPathSelectElement(string.Format("one:Notebook[@name=\"{0}\"]", notebookName), hierarchy.Xnm);
            if (bibleNotebook == null)
                bibleNotebook = hierarchy.Content.Root.XPathSelectElement(string.Format("one:Notebook[@nickname=\"{0}\"]", notebookName), hierarchy.Xnm);
            if (bibleNotebook != null)
            {
                return (string)bibleNotebook.Attribute("ID");
            }

            return string.Empty;
        }

        public static string GetNotebookIdByPath(ref Application oneNoteApp, string localPath, bool refreshCache, out string notebookName)
        {
            notebookName = null;
            var hierarchy = OneNoteProxy.Instance.GetHierarchy(ref oneNoteApp, null, HierarchyScope.hsNotebooks, refreshCache);
            var bibleNotebook = hierarchy.Content.Root.XPathSelectElement(string.Format("one:Notebook[@path=\"{0}\"]", localPath), hierarchy.Xnm);            
            if (bibleNotebook != null)
            {
                notebookName = (string)bibleNotebook.Attribute("name");
                return (string)bibleNotebook.Attribute("ID");
            }

            return string.Empty;
        }

        public static string GetNotebookElementNickname(ref Application oneNoteApp, string elementId)
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

        public static XDocument GetXDocument(string xml, out XmlNamespaceManager xnm, bool setLineInfo = false)
        {
            XDocument xd = !setLineInfo ? XDocument.Parse(xml) : XDocument.Parse(xml, LoadOptions.SetLineInfo);
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
                if (OneNoteUtils.IsError(ex, Error.hrObjectDoesNotExist))
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

            return parentEl.Root.XPathSelectElement(string.Format("one:{0}[@name=\"{1}\"]", elementTag, elementName), xnm);
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


        public static string GetOrGenerateLinkHref(ref Application oneNoteApp, string objectHref, string pageId, string objectId, params string[] additionalLinkQueryParameters)
        {
            string link;

            if (!string.IsNullOrEmpty(objectHref))
                link = objectHref;
            else
                link = OneNoteProxy.Instance.GenerateHref(ref oneNoteApp, pageId, objectId);

            foreach (var param in additionalLinkQueryParameters)
                link += "&" + param;

            return link;
        }

        public static string GetLink(string title, string link)
        {
            return string.Format("<a href=\"{0}\">{1}</a>", link, title);            
        }

        public static string GetOrGenerateLink(ref Application oneNoteApp, string title, string objectHref, string pageId, string objectId, params string[] additionalLinkQueryParameters)
        {
            var link = GetOrGenerateLinkHref(ref oneNoteApp, objectHref, pageId, objectId, additionalLinkQueryParameters);

            return GetLink(title, link);
        }

        public static string GenerateLink(ref Application oneNoteApp, string title, string pageId, string objectId, params string[] additionalLinkQueryParameters)
        {
            return GetOrGenerateLink(ref oneNoteApp, title, null, pageId, objectId, additionalLinkQueryParameters);
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


        public static void UpdatePageContentSafe(ref Application oneNoteApp, XDocument pageContent, XmlNamespaceManager xnm, bool repeatIfPageIsReadOnly = true)
        {
            UpdatePageContentSafeInternal(ref oneNoteApp, pageContent, xnm, repeatIfPageIsReadOnly ? (int?)0 : null);
        }

        private static void UpdatePageContentSafeInternal(ref Application oneNoteApp, XDocument pageContent, XmlNamespaceManager xnm, int? attemptsCount)
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
                try
                {
                UseOneNoteAPI(ref oneNoteApp, (oneNoteAppSafe) =>
                {
                    oneNoteAppSafe.UpdatePageContent(pageContent.ToString(), DateTime.MinValue, Constants.CurrentOneNoteSchema);
                });
            }
            catch (COMException ex)
            {
                    if (attemptsCount.GetValueOrDefault(int.MaxValue) < 20)  // 10 секунд - но каждое обновление требует времени. потому на самом деле дольше
                    {
                        if (OneNoteUtils.IsError(ex, Error.hrPageReadOnly) || OneNoteUtils.IsError(ex, Error.hrSectionReadOnly))
                        {
                            Thread.Sleep(500);
                            UpdatePageContentSafeInternal(ref oneNoteApp, pageContent, xnm, attemptsCount + 1);
                        }
                    }
                    else
                        throw;
                }
            }
            catch (COMException ex)
            {
                if (OneNoteUtils.IsError(ex, Error.hrInsertingInk))
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
                if (ex.Message.Contains("0x80010100")                                           // "System.Runtime.InteropServices.COMException (0x80010100): System call failed. (Exception from HRESULT: 0x80010100 (RPC_E_SYS_CALL_FAILED))";
                    || ex.Message.Contains("0x800706BA") 
                    || ex.Message.Contains("0x800706BE")
                    || ex.Message.Contains("0x80010001")                                        // System.Runtime.InteropServices.COMException (0x80010001): Вызов был отклонен. (Исключение из HRESULT: 0x80010001 (RPC_E_CALL_REJECTED))
                    )  
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

        public static void UpdateElementMetaData(XElement el, string key, string value, XmlNamespaceManager xnm)
        {
            var metaElement = el.XPathSelectElement(string.Format("one:Meta[@name=\"{0}\"]", key), xnm);
            if (metaElement != null)
            {
                metaElement.SetAttributeValue("content", value);
            }
            else
            {
                XNamespace nms = XNamespace.Get(Constants.OneNoteXmlNs);

                var pageSettings = el.XPathSelectElement("one:MediaPlaylist", xnm) ?? el.XPathSelectElement("one:PageSettings", xnm);
                
                var meta = new XElement(nms + "Meta",
                                            new XAttribute("name", key),
                                            new XAttribute("content", value));


                if (pageSettings != null)
                    pageSettings.AddBeforeSelf(meta);
                else
                    el.AddFirst(meta);
            }
        }


        public static string GetElementMetaData(XElement el, string key, XmlNamespaceManager xnm)
        {
            var metaElement = el.XPathSelectElement(string.Format("one:Meta[@name=\"{0}\"]", key), xnm);
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
                NotebookId = currentNotebookId,
                SectionGroupId = currentSectionGroupId,
                SectionId = currentSectionId,
                Id = currentPageId
            };
        }

        public static string GetElementPath(ref Application oneNoteApp, string elementId)
        {
            XmlNamespaceManager xnm;
            var xDoc = OneNoteUtils.GetHierarchyElement(ref oneNoteApp, elementId, HierarchyScope.hsSelf, out xnm);
            return (string)xDoc.Root.Attribute("path");
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="oneNoteApp"></param>
        /// <param name="notebookId"></param>
        /// <returns>false - если хранится в skydrive</returns>
        public static bool IsNotebookLocal(ref Application oneNoteApp, string notebookId)
        {
            try
            {
                string folderPath = GetElementPath(ref oneNoteApp, notebookId);

                Directory.GetCreationTime(folderPath);

                return true;
            }
            catch (NotSupportedException)
            {
                return false;
            }
        }     

        public static string ParseNotebookName(string s)
        {   
            if (!string.IsNullOrEmpty(s))
            {
                return s.Split(new string[] { OneNoteUtils.NotebookNameDelimeter }, StringSplitOptions.None)[0];
            }

            return s;
        }
        public static string NotebookNameDelimeter = " [\"";
        public static Dictionary<string, string> GetExistingNotebooks(ref Application oneNoteApp)
        {
            Dictionary<string, string> result = new Dictionary<string, string>();

            OneNoteProxy.HierarchyElement hierarchy = OneNoteProxy.Instance.GetHierarchy(ref oneNoteApp, null, HierarchyScope.hsNotebooks, true);

            foreach (XElement notebook in hierarchy.Content.Root.XPathSelectElements("one:Notebook", hierarchy.Xnm))
            {
                var name = (string)notebook.Attribute("nickname");
                var id = (string)notebook.Attribute("ID");

                result.Add(id, name);
            }

            return result;
        }

        public static bool IsError(Exception ex, Error error)
        {
            return ex.Message.IndexOf(error.ToString(), StringComparison.InvariantCultureIgnoreCase) > -1
                || ex.Message.IndexOf(GetHexError(error), StringComparison.InvariantCultureIgnoreCase) > -1;                
    }

        private static string GetHexError(Error error)
        {
            return string.Format("0x{0}", Convert.ToString((int)error, 16));
}

        public static string ParseError(string exceptionMessage)
        {
            var originalHexValue = Regex.Match(exceptionMessage, @"0x800[A-F\d]+").Value;
            if (!string.IsNullOrEmpty(originalHexValue))
            {
                var hexValue = originalHexValue.Replace("0x", "FFFFFFFF");
                long decValue;
                if (long.TryParse(hexValue, System.Globalization.NumberStyles.HexNumber, CultureInfo.CurrentCulture.NumberFormat, out decValue))
                    exceptionMessage = exceptionMessage.Replace(originalHexValue, ((Error)decValue).ToString());
            }

            return exceptionMessage;
        }
    }
}
