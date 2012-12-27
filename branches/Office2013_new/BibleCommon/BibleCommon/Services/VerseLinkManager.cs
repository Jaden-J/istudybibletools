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
        /// <param name="isSummaryNotesPage">Найти (создать) страницу, которая является страницей сводной заметок. Если false - то комментарий к Библии</param>
        /// <param name="pageLevel">1, 2 or 3</param>
        /// <returns>target pageId</returns>
        public static string FindVerseLinkPageAndCreateIfNeeded(ref Application oneNoteApp, 
            string bibleSectionId, string biblePageId, string biblePageName, string descriptionPageName,
            bool isSummaryNotesPage, out bool pageWasCreated, string verseLinkParentPageId = null, int pageLevel = 1, bool createIfNeeded = true)
        {
            string result = string.Empty;
            pageWasCreated = false;

            string exceptionResolveWay = isSummaryNotesPage ? string.Empty : string.Format("\n{0}", BibleCommon.Resources.Constants.VerseLinkManagerOpenNotBiblePage);
            string sectionGroupId = FindDescriptionSectionGroupForBiblePage(ref oneNoteApp, bibleSectionId, createIfNeeded, isSummaryNotesPage, false);
            if (!string.IsNullOrEmpty(sectionGroupId))
            {
                string sectionId = FindDescriptionSectionForBiblePage(ref oneNoteApp, biblePageName, sectionGroupId, createIfNeeded, false);
                if (!string.IsNullOrEmpty(sectionId))
                {
                    string bibleSectionName = OneNoteUtils.GetHierarchyElementName(ref oneNoteApp, bibleSectionId);
                    string pageId = FindDescriptionPageForBiblePage(ref oneNoteApp, sectionId,
                        bibleSectionName, biblePageId, biblePageName, descriptionPageName, isSummaryNotesPage, verseLinkParentPageId, pageLevel, createIfNeeded, out pageWasCreated);
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


        private static string FindDescriptionSectionGroupForBiblePage(ref Application oneNoteApp,
            string bibleSectionId, bool createIfNeeded, bool isSummaryNotesPage, bool refreshCache)
        {
            OneNoteProxy.HierarchyElement bibleDocument = OneNoteProxy.Instance.GetHierarchy(ref oneNoteApp, SettingsManager.Instance.NotebookId_Bible, HierarchyScope.hsSections, refreshCache);
            OneNoteProxy.HierarchyElement commentsDocument = OneNoteProxy.Instance.GetHierarchy(ref oneNoteApp, 
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
                        commentsDocument.Xnm);                                        // Комментарии к Библии/Новый Завет

                if (targetParentSectionGroup != null)
                {
                    XElement targetSectionGroup = targetParentSectionGroup.XPathSelectElement(
                        string.Format("one:SectionGroup[@name='{0}']", sectionName), commentsDocument.Xnm);             // Комментарии к Библии/Новый Завет/01. От Матфея

                    if (targetSectionGroup == null)
                    {
                        if (createIfNeeded)
                        {
                            CreateDescriptionSectionGroupForBiblePage(ref oneNoteApp, commentsDocument.Content, targetParentSectionGroup, sectionName);

                            return FindDescriptionSectionGroupForBiblePage(ref oneNoteApp, bibleSectionId, createIfNeeded, isSummaryNotesPage, true);  // надо обновить XML
                        }
                    }
                    else
                        return (string)targetSectionGroup.Attribute("ID");
                }

            }

            return string.Empty;
        }

        private static void CreateDescriptionSectionGroupForBiblePage(ref Application oneNoteApp,
            XDocument document, XElement targetParentSectionGroup, string sectionName)
        {
            XNamespace nms = XNamespace.Get(Constants.OneNoteXmlNs);
            XElement targetSectionGroup = new XElement(nms + "SectionGroup",
                                    new XAttribute("name", sectionName));

            targetParentSectionGroup.Add(targetSectionGroup);

            OneNoteUtils.UseOneNoteAPI(ref oneNoteApp, (oneNoteAppSafe) =>
            {
                oneNoteAppSafe.UpdateHierarchy(document.ToString(), Constants.CurrentOneNoteSchema);
            });

            OneNoteUtils.UseOneNoteAPI(ref oneNoteApp, (oneNoteAppSafe) =>
            {
                oneNoteAppSafe.SyncHierarchy((string)targetParentSectionGroup.Attribute("ID"));
            });
        }

        private static string FindDescriptionSectionForBiblePage(ref Application oneNoteApp,
            string biblePageName, string targetSectionGroupId, bool createIfNeeded, bool refreshCache)
        {
            OneNoteProxy.HierarchyElement sectionGroupDocument = OneNoteProxy.Instance.GetHierarchy(ref oneNoteApp, targetSectionGroupId, HierarchyScope.hsSections, refreshCache);            
            string sectionGroupName = (string)sectionGroupDocument.Content.Root.Attribute("name");

            if (sectionGroupName.IndexOf(biblePageName) != -1)
                biblePageName = SettingsManager.Instance.SectionName_DefaultBookOverview;

            XElement targetSection = sectionGroupDocument.Content.Root.XPathSelectElement(
                string.Format("one:Section[@name='{0}']", biblePageName), sectionGroupDocument.Xnm);

            if (targetSection == null)
            {
                if (createIfNeeded)
                {
                    CreateDescriptionSectionForBiblePage(ref oneNoteApp, sectionGroupDocument.Content, biblePageName);

                    return FindDescriptionSectionForBiblePage(ref oneNoteApp, biblePageName, targetSectionGroupId, createIfNeeded, true);  // надо обновить XML
                }
            }
            else
                return (string)targetSection.Attribute("ID");

            return string.Empty;
        }

        private static void CreateDescriptionSectionForBiblePage(ref Application oneNoteApp, XDocument sectionGroupDocument, string pageName)
        {
            XNamespace nms = XNamespace.Get(Constants.OneNoteXmlNs);
            XElement targetSection = new XElement(nms + "Section",
                                    new XAttribute("name", pageName));

            if (pageName == SettingsManager.Instance.SectionName_DefaultBookOverview || sectionGroupDocument.Root.Nodes().Count() == 0)
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

            OneNoteUtils.UseOneNoteAPI(ref oneNoteApp, (oneNoteAppSafe) =>
            {
                oneNoteAppSafe.UpdateHierarchy(sectionGroupDocument.ToString(), Constants.CurrentOneNoteSchema);                
            });

            OneNoteUtils.UseOneNoteAPI(ref oneNoteApp, (oneNoteAppSafe) =>
            {
                oneNoteAppSafe.SyncHierarchy((string)sectionGroupDocument.Root.Attribute("ID"));
            });
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
        private static string FindDescriptionPageForBiblePage(ref Application oneNoteApp, string sectionId, 
            string bibleSectionName, string biblePageId, string biblePageName, string descriptionPageName,
            bool isSummaryNotesPage, string verseLinkParentPageId, int pageLevel, bool createIfNeeded, out bool pageWasCreated)
        {
            pageWasCreated = false;
            OneNoteProxy.HierarchyElement sectionDocument = OneNoteProxy.Instance.GetHierarchy(ref oneNoteApp, sectionId, HierarchyScope.hsPages);            

            VersePointer vp = GetVersePointer(bibleSectionName, biblePageName);

            //if (vp == null)
            //    throw new Exception(
            //        string.Format(
            //            "Не удаётся найти соответствующее место Писания для страницы. bibleSectionName = '{0}', biblePageName = '{1}'",
            //            bibleSectionName, biblePageName));

            string pageDisplayName = string.Format("{0}.{1}", descriptionPageName, 
                                                               (vp != null && vp.IsValid) ? string.Format(" [{0}]", vp.ChapterName) : string.Empty);

            XElement page = sectionDocument.Content.Root.XPathSelectElement(string.Format("one:Page[@name='{0}']", pageDisplayName), sectionDocument.Xnm);

            string pageId = string.Empty;

            if (page == null)
            {
                if (createIfNeeded)
                {
                    OneNoteUtils.UseOneNoteAPI(ref oneNoteApp, (oneNoteAppSafe) =>
                    {
                        oneNoteAppSafe.CreateNewPage(sectionId, out pageId, NewPageStyle.npsBlankPageWithTitle);
                    });
                    pageWasCreated = true;

                    OneNoteProxy.PageContent biblePageDoc = OneNoteProxy.Instance.GetPageContent(ref oneNoteApp, biblePageId, OneNoteProxy.PageType.Bible);
                    string biblePageTitleId = (string)biblePageDoc.Content.Root
                        .XPathSelectElement("one:Title/one:OE", biblePageDoc.Xnm).Attribute("objectID");

                    string pageName = descriptionPageName + ".";

                    if (vp != null)
                    {   
                        string linkToBiblePage = OneNoteUtils.GetOrGenerateLink(ref oneNoteApp, vp.ChapterName, null, biblePageId, biblePageTitleId,
                                                                    Consts.Constants.QueryParameter_BibleVerse, Consts.Constants.QueryParameter_QuickAnalyze);
                        pageName += string.Format(" <span style='font-size:10pt;'>[{0}]</span>", linkToBiblePage);
                    }

                    var pageEl = SetPageName(ref oneNoteApp, pageId, pageName, isSummaryNotesPage, pageLevel);              

                    OneNoteProxy.Instance.RegisterVerseLinkSortPage(sectionId, pageId, verseLinkParentPageId, pageLevel);         
           
                    var sectionHierarchyCache = OneNoteProxy.Instance.GetHierarchy(ref oneNoteApp, sectionId, HierarchyScope.hsPages);
                    sectionHierarchyCache.Content.Root.Add(pageEl); // обновляем кэш иерархии
                }
            }
            else
                pageId = (string)page.Attribute("ID");

            return pageId;
        }

        public static void SortVerseLinkPages(ref Application oneNoteApp, string sectionId, string newPageId, string verseLinkParentPageId, int pageLevel)
        {
            OneNoteProxy.HierarchyElement hierarchy = OneNoteProxy.Instance.GetHierarchy(ref oneNoteApp, sectionId, HierarchyScope.hsPages);
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
        private static XElement SetPageName(ref Application oneNoteApp, string pageId, string pageName, bool isSummaryNotesPage, int pageLevel)
        {
            XNamespace nms = XNamespace.Get(Constants.OneNoteXmlNs);

            var pageContent = OneNoteProxy.Instance.GetPageContent(ref oneNoteApp, pageId, isSummaryNotesPage ? OneNoteProxy.PageType.NotesPage : OneNoteProxy.PageType.CommentPage);
            pageContent.Content.Root.SetAttributeValue("pageLevel", pageLevel);
            var title = pageContent.Content.Root.XPathSelectElement("one:Title/one:OE/one:T", pageContent.Xnm);
            if (title != null)
                title.Value = pageName;

            var pageEl = new XElement(nms + "Page",
                            new XAttribute("ID", pageId),
                            new XAttribute("name", pageName),
                            new XAttribute("pageLevel", pageLevel),
                            new XAttribute("dateTime", pageContent.Content.Root.Attribute("dateTime").Value),
                            new XAttribute("lastModifiedTime", pageContent.Content.Root.Attribute("lastModifiedTime").Value)
                        );

            if (isSummaryNotesPage)
            {
                OneNoteUtils.UpdateElementMetaData(pageContent.Content.Root, Constants.Key_IsSummaryNotesPage, isSummaryNotesPage.ToString(), pageContent.Xnm);
                OneNoteUtils.UpdateElementMetaData(pageEl, Constants.Key_IsSummaryNotesPage, isSummaryNotesPage.ToString(), pageContent.Xnm);

                OneNoteUtils.UpdateElementMetaData(pageContent.Content.Root, Constants.Key_NotesPageManagerName, NotesPageManagerEx.Const_ManagerName, pageContent.Xnm);
                OneNoteUtils.UpdateElementMetaData(pageEl, Constants.Key_NotesPageManagerName, NotesPageManagerEx.Const_ManagerName, pageContent.Xnm);
            }

            pageContent.WasModified = true;

            return pageEl;
        }
    }
}
