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
        /// <summary>
        /// 
        /// </summary>
        /// <param name="oneNoteApp"></param>
        /// <param name="currentSectionId"></param>
        /// <param name="currentPageId"></param>
        /// <param name="currentPageName"></param>
        /// <param name="descriptionPageName"></param>
        /// <returns>target pageId</returns>
        public static string FindVerseLinkPageAndCreateIfNeeded(Application oneNoteApp, 
            string currentSectionId, string currentPageId, string currentPageName, string descriptionPageName, bool createIfNeeded = true)
        {

            string sectionGroupId = FindDescriptionSectionGroupForCurrentPage(oneNoteApp, currentSectionId, createIfNeeded);
            if (!string.IsNullOrEmpty(sectionGroupId))
            {
                string sectionId = FindDescriptionSectionForCurrentPage(oneNoteApp, currentPageName, sectionGroupId, createIfNeeded);
                if (!string.IsNullOrEmpty(sectionId))
                {
                    string currentSectionName = OneNoteUtils.GetHierarchyElementName(oneNoteApp, currentSectionId);
                    string pageId = FindDescriptionPageForCurrentPage(oneNoteApp, sectionId,
                        currentSectionName, currentPageId, currentPageName, descriptionPageName, createIfNeeded);
                    if (!string.IsNullOrEmpty(pageId))
                    {
                        return pageId;
                    }
                    else
                        throw new NotFoundVerseLinkPageExceptions("Не найдена страница для комментариев");
                }
                else
                    throw new NotFoundVerseLinkPageExceptions("Не найдена секция для комментариев");
            }
            else
                throw new NotFoundVerseLinkPageExceptions("Не найдена группа секций для комментариев");
        }


        private static string FindDescriptionSectionGroupForCurrentPage(Application oneNoteApp,
            string currentSectionId, bool createIfNeeded, bool refreshCache = false)
        {
            OneNoteProxy.HierarchyElement bibleDocument = OneNoteProxy.Instance.GetHierarchy(oneNoteApp, SettingsManager.Instance.NotebookId_Bible, HierarchyScope.hsSections, refreshCache);
            OneNoteProxy.HierarchyElement commentsDocument = OneNoteProxy.Instance.GetHierarchy(oneNoteApp, SettingsManager.Instance.NotebookId_BibleComments, HierarchyScope.hsSections, refreshCache);             

            XElement currentSection = bibleDocument.Content.Root.XPathSelectElement(string.Format("{0}one:SectionGroup/one:Section[@ID='{1}']",
                !string.IsNullOrEmpty(SettingsManager.Instance.SectionGroupId_Bible) ? "one:SectionGroup/" : string.Empty, currentSectionId), bibleDocument.Xnm);

            if (currentSection != null && currentSection.Parent != null && currentSection.Parent.Parent != null)
            {
                string sectionName = (string)currentSection.Attribute("name");                          // 01. От Матфея
                string sectionGroupName = (string)currentSection.Parent.Attribute("name");              // Новый Завет                

                XElement targetParentSectionGroup = commentsDocument.Content.Root.XPathSelectElement(
                        string.Format("{0}one:SectionGroup[@name='{1}']",                    
                            !string.IsNullOrEmpty(SettingsManager.Instance.SectionGroupId_BibleComments) 
                                ? string.Format("one:SectionGroup[@ID='{0}']/", SettingsManager.Instance.SectionGroupId_BibleComments)
                                : string.Empty,
                            sectionGroupName), 
                        commentsDocument.Xnm);                                        // Изучение Библии/Новый Завет

                if (targetParentSectionGroup != null)
                {
                    XElement targetSectionGroup = targetParentSectionGroup.XPathSelectElement(
                        string.Format("one:SectionGroup[@name='{0}']", sectionName), commentsDocument.Xnm);             // Изучение Библии/Новый Завет/01. От Матфея

                    if (targetSectionGroup == null)
                    {
                        if (createIfNeeded)
                        {
                            CreateDescriptionSectionGroupForCurrentPage(oneNoteApp, commentsDocument.Content, targetParentSectionGroup, sectionName);

                            return FindDescriptionSectionGroupForCurrentPage(oneNoteApp, currentSectionId, createIfNeeded, true);  // надо обновить XML
                        }
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

        private static string FindDescriptionSectionForCurrentPage(Application oneNoteApp,
            string currentPageName, string targetSectionGroupId, bool createIfNeeded, bool refreshCache = false)
        {
            OneNoteProxy.HierarchyElement sectionGroupDocument = OneNoteProxy.Instance.GetHierarchy(oneNoteApp, targetSectionGroupId, HierarchyScope.hsSections, refreshCache);            
            string sectionGroupName = (string)sectionGroupDocument.Content.Root.Attribute("name");

            if (sectionGroupName.IndexOf(currentPageName) != -1)
                currentPageName = SettingsManager.Instance.PageName_DefaultBookOverview;

            XElement targetSection = sectionGroupDocument.Content.Root.XPathSelectElement(
                string.Format("one:Section[@name='{0}']", currentPageName), sectionGroupDocument.Xnm);

            if (targetSection == null)
            {
                if (createIfNeeded)
                {
                    CreateDescriptionSectionForCurrentPage(oneNoteApp, sectionGroupDocument.Content, currentPageName);

                    return FindDescriptionSectionForCurrentPage(oneNoteApp, currentPageName, targetSectionGroupId, createIfNeeded, true);  // надо обновить XML
                }
            }
            else
                return (string)targetSection.Attribute("ID");

            return string.Empty;
        }

        private static void CreateDescriptionSectionForCurrentPage(Application oneNoteApp, XDocument sectionGroupDocument, string pageName)
        {
            XNamespace nms = XNamespace.Get(Constants.OneNoteXmlNs);
            XElement targetSection = new XElement(nms + "Section",
                                    new XAttribute("name", pageName));

            if (pageName == SettingsManager.Instance.PageName_DefaultBookOverview || sectionGroupDocument.Root.Nodes().Count() == 0)
                sectionGroupDocument.Root.AddFirst(targetSection);
            else
            {
                int? pageNameIndex = StringUtils.GetStringFirstNumber(pageName);
                bool wasAdded = false;
                foreach (XElement section in sectionGroupDocument.Root.Nodes())
                {
                    string name = (string)section.Attribute("name");
                    int? otherPageIndex = StringUtils.GetStringFirstNumber(name);

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

        private static string FindDescriptionPageForCurrentPage(Application oneNoteApp, string sectionId, 
            string currentSectionName, string currentPageId, string currentPageName, string descriptionPageName, bool createIfNeeded)
        {   
            OneNoteProxy.HierarchyElement sectionDocument = OneNoteProxy.Instance.GetHierarchy(oneNoteApp, sectionId, HierarchyScope.hsPages);            

            VersePointer vp = GetCurrentVersePointer(currentSectionName, currentPageName);

            if (vp == null)
                throw new Exception(
                    string.Format(
                        "Не удаётся найти соответствующее место Писания для текущей страницы. currentSectionName = '{0}', currentPageName = '{1}'",
                        currentSectionName, currentPageName));

            string pageDisplayName = string.Format("{0}. [{1}]", descriptionPageName, vp.FriendlyChapterName);

            XElement page = sectionDocument.Content.Root.XPathSelectElement(string.Format("one:Page[@name='{0}']", pageDisplayName), sectionDocument.Xnm);

            string pageId = string.Empty;

            if (page == null)
            {
                if (createIfNeeded)
                {
                    oneNoteApp.CreateNewPage(sectionId, out pageId, NewPageStyle.npsBlankPageWithTitle);

                    OneNoteProxy.PageContent currentPageDoc = OneNoteProxy.Instance.GetPageContent(oneNoteApp, currentPageId, OneNoteProxy.PageType.Bible);
                    string currentPageTitleId = (string)currentPageDoc.Content.Root
                        .XPathSelectElement("one:Title/one:OE", currentPageDoc.Xnm).Attribute("objectID");

                    string linkToCurrentPage = OneNoteUtils.GenerateHref(oneNoteApp, vp.FriendlyChapterName, currentPageId, currentPageTitleId);
                    
                    string pageName = string.Format("{0}. <span style='font-size:10pt;'>[{1}]</span>",
                                        descriptionPageName, linkToCurrentPage);
                    SetPageName(oneNoteApp, pageId, pageName);

                    OneNoteProxy.Instance.RefreshHierarchyCache(oneNoteApp, sectionId, HierarchyScope.hsPages);
                }
            }
            else
                pageId = (string)page.Attribute("ID");

            return pageId;
        }

        private static VersePointer GetCurrentVersePointer(string currentSectionName, string currentPageName)
        {
            VersePointer result = null;

            int? chapter = StringUtils.GetStringFirstNumber(currentPageName);
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
