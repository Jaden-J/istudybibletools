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
        /// <param name="bibleSectionId"></param>
        /// <param name="biblePageId"></param>
        /// <param name="biblePageName"></param>
        /// <param name="descriptionPageName"></param>
        /// <param name="isSummaryNotesPage">Найти (создать) страницу, которая является страницей сводной заметок?</param>
        /// <param name="pageLevel">1, 2 or 3</param>
        /// <returns>target pageId</returns>
        public static string FindVerseLinkPageAndCreateIfNeeded(Application oneNoteApp, 
            string bibleSectionId, string biblePageId, string biblePageName, string descriptionPageName,
            bool isSummaryNotesPage, string verseLinkParentPageId = null, int pageLevel = 1, bool createIfNeeded = true)
        {
            string result = string.Empty;

            string exceptionResolveWay = string.Format("\n{0}", BibleCommon.Resources.Constants.VerseLinkManagerOpenNotBiblePage);
            string sectionGroupId = FindDescriptionSectionGroupForBiblePage(oneNoteApp, bibleSectionId, createIfNeeded, isSummaryNotesPage, false);
            if (!string.IsNullOrEmpty(sectionGroupId))
            {
                string sectionId = FindDescriptionSectionForBiblePage(oneNoteApp, biblePageName, sectionGroupId, createIfNeeded, false);
                if (!string.IsNullOrEmpty(sectionId))
                {
                    string bibleSectionName = OneNoteUtils.GetHierarchyElementName(oneNoteApp, bibleSectionId);
                    string pageId = FindDescriptionPageForBiblePage(oneNoteApp, sectionId,
                        bibleSectionName, biblePageId, biblePageName, descriptionPageName, isSummaryNotesPage, verseLinkParentPageId, pageLevel, createIfNeeded);
                    if (!string.IsNullOrEmpty(pageId))
                    {
                        result = pageId;
                    }
                    else if (createIfNeeded)
                        throw new NotFoundVerseLinkPageExceptions(BibleCommon.Resources.Constants.VerseLinkManagerCommentPageNotFound + exceptionResolveWay);
                }
                else if (createIfNeeded)
                    throw new NotFoundVerseLinkPageExceptions(BibleCommon.Resources.Constants.VerseLinkManagerCommentSectionNotFound + exceptionResolveWay);
            }
            else if (createIfNeeded)
                throw new NotFoundVerseLinkPageExceptions(BibleCommon.Resources.Constants.VerseLinkManagerCommentSectionGroupNotFound + exceptionResolveWay);

            return result;
        }


        private static string FindDescriptionSectionGroupForBiblePage(Application oneNoteApp,
            string bibleSectionId, bool createIfNeeded, bool isSummaryNotesPage, bool refreshCache)
        {
            OneNoteProxy.HierarchyElement bibleDocument = OneNoteProxy.Instance.GetHierarchy(oneNoteApp, SettingsManager.Instance.NotebookId_Bible, HierarchyScope.hsSections, refreshCache);
            OneNoteProxy.HierarchyElement commentsDocument = OneNoteProxy.Instance.GetHierarchy(oneNoteApp, 
                isSummaryNotesPage ? SettingsManager.Instance.NotebookId_BibleNotesPages : SettingsManager.Instance.NotebookId_BibleComments, 
                HierarchyScope.hsSections, refreshCache);             

            XElement bibleSection = bibleDocument.Content.Root.XPathSelectElement(string.Format("{0}one:SectionGroup/one:Section[@ID='{1}']",
                !string.IsNullOrEmpty(SettingsManager.Instance.SectionGroupId_Bible) ? "one:SectionGroup/" : string.Empty, bibleSectionId), bibleDocument.Xnm);

            if (bibleSection != null && bibleSection.Parent != null && bibleSection.Parent.Parent != null)
            {
                string sectionName = (string)bibleSection.Attribute("name");                          // 01. От Матфея
                string sectionGroupName = (string)bibleSection.Parent.Attribute("name");              // Новый Завет                


                string rootSectionGroupId = isSummaryNotesPage ? SettingsManager.Instance.SectionGroupId_BibleNotesPages 
                                                        : SettingsManager.Instance.SectionGroupId_BibleComments;

                XElement targetParentSectionGroup = commentsDocument.Content.Root.XPathSelectElement(
                        string.Format("{0}one:SectionGroup[@name='{1}']",                    
                            !string.IsNullOrEmpty(rootSectionGroupId) 
                                ? string.Format("one:SectionGroup[@ID='{0}']/", rootSectionGroupId)
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
                            CreateDescriptionSectionGroupForBiblePage(oneNoteApp, commentsDocument.Content, targetParentSectionGroup, sectionName);

                            return FindDescriptionSectionGroupForBiblePage(oneNoteApp, bibleSectionId, createIfNeeded, isSummaryNotesPage, true);  // надо обновить XML
                        }
                    }
                    else
                        return (string)targetSectionGroup.Attribute("ID");
                }

            }

            return string.Empty;
        }

        private static void CreateDescriptionSectionGroupForBiblePage(Application oneNoteApp,
            XDocument document, XElement targetParentSectionGroup, string sectionName)
        {
            XNamespace nms = XNamespace.Get(Constants.OneNoteXmlNs);
            XElement targetSectionGroup = new XElement(nms + "SectionGroup",
                                    new XAttribute("name", sectionName));

            targetParentSectionGroup.Add(targetSectionGroup);

            oneNoteApp.UpdateHierarchy(document.ToString());
        }

        private static string FindDescriptionSectionForBiblePage(Application oneNoteApp,
            string biblePageName, string targetSectionGroupId, bool createIfNeeded, bool refreshCache)
        {
            OneNoteProxy.HierarchyElement sectionGroupDocument = OneNoteProxy.Instance.GetHierarchy(oneNoteApp, targetSectionGroupId, HierarchyScope.hsSections, refreshCache);            
            string sectionGroupName = (string)sectionGroupDocument.Content.Root.Attribute("name");

            if (sectionGroupName.IndexOf(biblePageName) != -1)
                biblePageName = SettingsManager.Instance.PageName_DefaultBookOverview;

            XElement targetSection = sectionGroupDocument.Content.Root.XPathSelectElement(
                string.Format("one:Section[@name='{0}']", biblePageName), sectionGroupDocument.Xnm);

            if (targetSection == null)
            {
                if (createIfNeeded)
                {
                    CreateDescriptionSectionForBiblePage(oneNoteApp, sectionGroupDocument.Content, biblePageName);

                    return FindDescriptionSectionForBiblePage(oneNoteApp, biblePageName, targetSectionGroupId, createIfNeeded, true);  // надо обновить XML
                }
            }
            else
                return (string)targetSection.Attribute("ID");

            return string.Empty;
        }

        private static void CreateDescriptionSectionForBiblePage(Application oneNoteApp, XDocument sectionGroupDocument, string pageName)
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="oneNoteApp"></param>
        /// <param name="sectionId"></param>
        /// <param name="bibleSectionName"></param>
        /// <param name="biblePageId"></param>
        /// <param name="biblePageName"></param>
        /// <param name="descriptionPageName"></param>
        /// <param name="pageLevel">1, 2 or 3</param>
        /// <param name="createIfNeeded"></param>
        /// <returns></returns>
        private static string FindDescriptionPageForBiblePage(Application oneNoteApp, string sectionId, 
            string bibleSectionName, string biblePageId, string biblePageName, string descriptionPageName,
            bool isSummaryNotesPage, string verseLinkParentPageId, int pageLevel, bool createIfNeeded)
        {   
            OneNoteProxy.HierarchyElement sectionDocument = OneNoteProxy.Instance.GetHierarchy(oneNoteApp, sectionId, HierarchyScope.hsPages);            

            VersePointer vp = GetVersePointer(bibleSectionName, biblePageName);

            //if (vp == null)
            //    throw new Exception(
            //        string.Format(
            //            "Не удаётся найти соответствующее место Писания для страницы. bibleSectionName = '{0}', biblePageName = '{1}'",
            //            bibleSectionName, biblePageName));

            string pageDisplayName = string.Format("{0}.{1}", descriptionPageName, 
                                                               vp != null ? string.Format(" [{0}]", vp.ChapterName) : string.Empty);

            XElement page = sectionDocument.Content.Root.XPathSelectElement(string.Format("one:Page[@name='{0}']", pageDisplayName), sectionDocument.Xnm);

            string pageId = string.Empty;

            if (page == null)
            {
                if (createIfNeeded)
                {
                    oneNoteApp.CreateNewPage(sectionId, out pageId, NewPageStyle.npsBlankPageWithTitle);

                    OneNoteProxy.PageContent biblePageDoc = OneNoteProxy.Instance.GetPageContent(oneNoteApp, biblePageId, OneNoteProxy.PageType.Bible);
                    string biblePageTitleId = (string)biblePageDoc.Content.Root
                        .XPathSelectElement("one:Title/one:OE", biblePageDoc.Xnm).Attribute("objectID");

                    string pageName = descriptionPageName + ".";

                    if (vp != null)
                    {
                        string linkToBiblePage = OneNoteUtils.GenerateHref(oneNoteApp, vp.ChapterName, biblePageId, biblePageTitleId);
                        pageName += string.Format(" <span style='font-size:10pt;'>[{0}]</span>", linkToBiblePage);
                    }

                    SetPageName(oneNoteApp, pageId, pageName, isSummaryNotesPage, pageLevel, biblePageDoc.Xnm);              

                    OneNoteProxy.Instance.RegisteVerseLinkSortPage(sectionId, pageId, verseLinkParentPageId, pageLevel);

                    OneNoteProxy.Instance.RefreshHierarchyCache(oneNoteApp, sectionId, HierarchyScope.hsPages);                    
                }
            }
            else
                pageId = (string)page.Attribute("ID");

            return pageId;
        }

        public static void SortVerseLinkPages(Application oneNoteApp, string sectionId, string newPageId, string verseLinkParentPageId, int pageLevel)
        {
            OneNoteProxy.HierarchyElement hierarchy = OneNoteProxy.Instance.GetHierarchy(oneNoteApp, sectionId, HierarchyScope.hsPages);
            var newPage = hierarchy.Content.Root.XPathSelectElement(string.Format("one:Page[@ID='{0}']", newPageId), hierarchy.Xnm);
            string newPageName = (string)newPage.Attribute("name");

            XElement prevPage = null;
            foreach (XElement oldPage in hierarchy.Content.Root.XPathSelectElements(string.Format("one:Page[@ID != '{0}']", newPageId), hierarchy.Xnm))
            {
                if (prevPage == null)
                {
                    if (!string.IsNullOrEmpty(verseLinkParentPageId))
                    {
                        string oldPageId = (string)oldPage.Attribute("ID");

                        if (oldPageId == verseLinkParentPageId)
                        {
                            prevPage = oldPage;
                            continue;
                        }
                        else
                            continue;
                    }
                }

                string oldPageLevel = (string)oldPage.Attribute("pageLevel");
                if (pageLevel <= int.Parse(oldPageLevel))
                {
                    string oldPageName = (string)oldPage.Attribute("name");

                    if (StringUtils.CompareTo(newPageName, oldPageName) < 0)
                        break;

                    prevPage = oldPage;
                }
                else
                    if (prevPage != null)  // если уже нашли страницу, после которой вставляем
                        break;
            }

            hierarchy.Content.Root.XPathSelectElement(string.Format("one:Page[@ID='{0}']", newPageId), hierarchy.Xnm).Remove();
            if (prevPage == null)
            {
                hierarchy.Content.Root.AddFirst(newPage);
            }
            else
            {
                prevPage.AddAfterSelf(newPage);
            }

            hierarchy.WasModified = true;

            //oneNoteApp.UpdateHierarchy(hierarchy.Content.ToString());

            //OneNoteProxy.Instance.RefreshHierarchyCache(oneNoteApp, sectionId, HierarchyScope.hsPages);
        }

        private static VersePointer GetVersePointer(string bibleSectionName, string biblePageName)
        {
            VersePointer result = null;

            int? chapter = StringUtils.GetStringFirstNumber(biblePageName);
            if (chapter.HasValue)
            {
                result = VersePointer.GetChapterVersePointer(string.Format("{0} {1}", bibleSectionName.Substring(4), chapter));
            }

            return result;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="oneNoteApp"></param>
        /// <param name="pageId"></param>
        /// <param name="pageName"></param>
        /// <param name="pageLevel">1, 2 or 3</param>
        private static void SetPageName(Application oneNoteApp, string pageId, string pageName, bool isSummaryNotesPage, int pageLevel, XmlNamespaceManager xnm)
        {
            XNamespace nms = XNamespace.Get(Constants.OneNoteXmlNs);
            XDocument pageDocument = new XDocument(new XElement(nms + "Page",                            
                            new XAttribute("ID", pageId),
                            new XAttribute("pageLevel", pageLevel),
                            new XElement(nms + "Title",
                                new XElement(nms + "OE",
                                    new XElement(nms + "T",
                                        new XCData(
                                            pageName
                                            ))))));

            if (isSummaryNotesPage)
                OneNoteUtils.UpdatePageMetaData(oneNoteApp, pageDocument.Root, Constants.Key_IsSummaryNotesPage, isSummaryNotesPage.ToString(), xnm);

            OneNoteUtils.UpdatePageContentSafe(oneNoteApp, pageDocument);
        }
    }
}
