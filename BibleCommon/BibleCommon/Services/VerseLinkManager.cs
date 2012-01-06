using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using Microsoft.Office.Interop.OneNote;
using System.Xml.Linq;
using System.Xml.XPath;
using BibleCommon.Helpers;
using BibleCommon.Consts;
using BibleCommon.Common;

namespace BibleCommon.Services
{
    public static class VerseLinkManager
    {
        public static string FindVerseLinkPageAndCreateIfNeeded(Application oneNoteApp, string currentNotebookId,
            string currentSectionId, string currentPageId, string currentPageName, string descriptionPageName)
        {
            string sectionGroupId = FindDescriptionSectionGroupForCurrentPage(oneNoteApp, currentNotebookId, currentSectionId);
            if (!string.IsNullOrEmpty(sectionGroupId))
            {
                string sectionId = FindDescriptionSectionForCurrentPage(oneNoteApp, currentPageName, sectionGroupId);
                if (!string.IsNullOrEmpty(sectionId))
                {
                    string currentSectionName = Utils.GetHierarchyElementName(oneNoteApp, currentSectionId);
                    string pageId = FindDescriptionPageForCurrentPage(oneNoteApp, sectionId,
                        currentSectionName, currentPageId, currentPageName, descriptionPageName);
                    if (!string.IsNullOrEmpty(pageId))
                    {
                        return pageId;
                    }
                    else
                        throw new Exception("Не найдена страница для комментариев");
                }
                else
                    throw new Exception("Не найдена секция для комментариев");
            }
            else
                throw new Exception("Не найдена группа секций для комментариев");
        }


        public static string FindDescriptionSectionGroupForCurrentPage(Application oneNoteApp,
            string currentNotebookId, string currentSectionId, bool refreshCache = false)
        {
            XmlNamespaceManager xnm;

            string notebookContentXml = OneNoteProxy.Instance.GetHierarchy(oneNoteApp, currentNotebookId, HierarchyScope.hsSections, refreshCache);

            XDocument document = Utils.GetXDocument(notebookContentXml, out xnm);

            XElement currentSection = document.XPathSelectElement(string.Format("/one:Notebook/one:SectionGroup/one:SectionGroup/one:Section[@ID='{0}']",
                currentSectionId), xnm);

            if (currentSection != null && currentSection.Parent != null && currentSection.Parent.Parent != null)
            {
                string sectionName = (string)currentSection.Attribute("name");                          // 01. От Матфея
                string sectionGroupName = (string)currentSection.Parent.Attribute("name");              // Новый Завет
                string topSectonGroupName = (string)currentSection.Parent.Parent.Attribute("name");     // Бибия

                XElement targetParentSectionGroup = document.XPathSelectElement(
                    string.Format("/one:Notebook/one:SectionGroup[@name!='{0}']/one:SectionGroup[@name='{1}']",
                    topSectonGroupName, sectionGroupName), xnm);                                        // Изучение Библии/Новый Завет

                if (targetParentSectionGroup != null)
                {
                    XElement targetSectionGroup = targetParentSectionGroup.XPathSelectElement(
                        string.Format("one:SectionGroup[@name='{0}']", sectionName), xnm);             // Изучение Библии/Новый Завет/01. От Матфея

                    if (targetSectionGroup == null)
                    {
                        CreateDescriptionSectionGroupForCurrentPage(oneNoteApp, document, targetParentSectionGroup, sectionName);

                        return FindDescriptionSectionGroupForCurrentPage(oneNoteApp, currentNotebookId, currentSectionId, true);  // надо обновить XML
                    }
                    else
                        return (string)targetSectionGroup.Attribute("ID");
                }

            }

            return string.Empty;
        }

        private static void CreateDescriptionSectionGroupForCurrentPage(Application oneNoteApp,
            XDocument document, XElement targetParentSectionGroup, string sectionName)
        {
            XNamespace nms = XNamespace.Get(Constants.OneNoteXmlNs);
            XElement targetSectionGroup = new XElement(nms + "SectionGroup",
                                    new XAttribute("name", sectionName));

            targetParentSectionGroup.Add(targetSectionGroup);

            oneNoteApp.UpdateHierarchy(document.ToString());
        }

        public static string FindDescriptionSectionForCurrentPage(Application oneNoteApp,
            string currentPageName, string targetSectionGroupId, bool refreshCache = false)
        {
            XmlNamespaceManager xnm;

            string sectionGroupXml = OneNoteProxy.Instance.GetHierarchy(oneNoteApp, targetSectionGroupId, HierarchyScope.hsSections, refreshCache);
            XDocument sectionGroupDocument = Utils.GetXDocument(sectionGroupXml, out xnm);
            string sectionGroupName = (string)sectionGroupDocument.Root.Attribute("name");

            if (sectionGroupName.IndexOf(currentPageName) != -1)
                currentPageName = SettingsManager.Instance.BookOverviewPageDefaultName;

            XElement targetSection = sectionGroupDocument.Root.XPathSelectElement(
                string.Format("one:Section[@name='{0}']", currentPageName), xnm);

            if (targetSection == null)
            {
                CreateDescriptionSectionForCurrentPage(oneNoteApp, sectionGroupDocument, currentPageName);

                return FindDescriptionSectionForCurrentPage(oneNoteApp, currentPageName, targetSectionGroupId, true);  // надо обновить XML
            }
            else
                return (string)targetSection.Attribute("ID");
        }

        private static void CreateDescriptionSectionForCurrentPage(Application oneNoteApp, XDocument sectionGroupDocument, string pageName)
        {
            XNamespace nms = XNamespace.Get(Constants.OneNoteXmlNs);
            XElement targetSection = new XElement(nms + "Section",
                                    new XAttribute("name", pageName));

            if (pageName == SettingsManager.Instance.BookOverviewPageDefaultName || sectionGroupDocument.Root.Nodes().Count() == 0)
                sectionGroupDocument.Root.AddFirst(targetSection);
            else
            {
                int? pageNameIndex = Utils.GetStringFirstNumber(pageName);
                bool wasAdded = false;
                foreach (XElement section in sectionGroupDocument.Root.Nodes())
                {
                    string name = (string)section.Attribute("name");
                    int? otherPageIndex = Utils.GetStringFirstNumber(name);

                    if (pageNameIndex.GetValueOrDefault(0) < otherPageIndex.GetValueOrDefault(0))
                    {
                        section.AddBeforeSelf(targetSection);
                        wasAdded = true;
                        break;
                    }
                }

                if (!wasAdded)
                    sectionGroupDocument.Root.Add(targetSection);
            }

            oneNoteApp.UpdateHierarchy(sectionGroupDocument.ToString());
        }

        public static string FindDescriptionPageForCurrentPage(Application oneNoteApp, string sectionId, 
            string currentSectionName, string currentPageId, string currentPageName, string descriptionPageName)
        {
            XmlNamespaceManager xnm;
            string sectionContentXml = OneNoteProxy.Instance.GetHierarchy(oneNoteApp, sectionId, HierarchyScope.hsPages);
            XDocument sectionDocument = Utils.GetXDocument(sectionContentXml, out xnm);

            VersePointer vp = GetCurrentVersePointer(currentSectionName, currentPageName);

            if (vp == null)
                throw new Exception(
                    string.Format(
                        "Не удаётся найти соответствующее место Писания для текущей страницы. currentSectionName = '{0}', currentPageName = '{1}'",
                        currentSectionName, currentPageName));

            string pageDisplayName = string.Format("{0}. [{1}]", descriptionPageName, vp.FriendlyChapterName);

            XElement page = sectionDocument.Root.XPathSelectElement(string.Format("one:Page[@name='{0}']", pageDisplayName), xnm);

            string pageId;

            if (page == null)
            {
                oneNoteApp.CreateNewPage(sectionId, out pageId, NewPageStyle.npsBlankPageWithTitle);

                string linkToCurrentPage;
                oneNoteApp.GetHyperlinkToObject(currentPageId, null, out linkToCurrentPage);
                string pageName = string.Format("{0}. <span style='font-size:10pt;'>[<a href='{1}'>{2}</a>]</span>", 
                                    descriptionPageName, linkToCurrentPage, vp.FriendlyChapterName);
                SetPageName(oneNoteApp, pageId, pageName);

                OneNoteProxy.Instance.RefreshHierarchyCache(oneNoteApp, sectionId, HierarchyScope.hsPages); 
            }
            else
                pageId = (string)page.Attribute("ID");

            return pageId;
        }

        private static VersePointer GetCurrentVersePointer(string currentSectionName, string currentPageName)
        {
            VersePointer result = null;
            
            int? chapter = Utils.GetStringFirstNumber(currentPageName);
            if (chapter.HasValue)
            {
                result = VersePointer.GetChapterVersePointer(string.Format("{0} {1}", currentSectionName.Substring(4), chapter));
            }

            return result;
        }

        private static void SetPageName(Application oneNoteApp, string pageId, string pageName)
        {
            XNamespace nms = XNamespace.Get(Constants.OneNoteXmlNs);
            XDocument pageDocument = new XDocument(new XElement(nms + "Page",
                            new XAttribute("ID", pageId),
                            new XElement(nms + "Title",
                                new XElement(nms + "OE",
                                    new XElement(nms + "T",
                                        new XCData(
                                            pageName
                                            ))))));

            oneNoteApp.UpdatePageContent(pageDocument.ToString());
        }
    }
}
