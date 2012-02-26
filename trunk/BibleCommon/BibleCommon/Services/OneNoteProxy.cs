using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.OneNote;
using System.Xml.Linq;
using BibleCommon.Helpers;
using System.Xml;

namespace BibleCommon.Services
{
    /// <summary>
    /// Кэш OneNote
    /// </summary>
    public class OneNoteProxy
    {
        public class OneNoteHierarchyContentId
        {
            public string ID { get; set; }
            public HierarchyScope ContentScope { get; set; }

            public override bool Equals(object obj)
            {
                return this.ID == ((OneNoteHierarchyContentId)obj).ID && this.ContentScope == ((OneNoteHierarchyContentId)obj).ContentScope;
            }

            public override int GetHashCode()
            {
                return this.ID.GetHashCode() ^ this.ContentScope.GetHashCode();
            }
        }

        public class BiblePageId
        {
            public string SectionId { get; set; }
            public string PageId { get; set; }
            public string PageName { get; set; }

            public override int GetHashCode()
            {
                return SectionId.GetHashCode() ^ PageId.GetHashCode() ^ PageName.GetHashCode();
            }

            public override bool Equals(object obj)
            {
                BiblePageId otherObject = (BiblePageId)obj;
                return SectionId == otherObject.SectionId
                    && PageId == otherObject.PageId
                    && PageName == otherObject.PageName;
            }
        }

        public class CommentPageId 
        {
            public BiblePageId BiblePageId { get; set; }
            public string CommentsPageName { get; set; }

            public override int GetHashCode()
            {
                return BiblePageId.GetHashCode() ^ CommentsPageName.GetHashCode();
            }

            public override bool Equals(object obj)
            {
                CommentPageId otherObject = (CommentPageId)obj;
                return BiblePageId.Equals(otherObject.BiblePageId)
                    && CommentsPageName == otherObject.CommentsPageName;
            }
        }

        public class PageContent
        {            
            public string PageId { get; set; }
            public XDocument Content { get; set; }
            public XmlNamespaceManager Xnm { get; set; }
            public bool WasModified { get; set; }
        }

        public class HierarchyElement
        {
            public OneNoteHierarchyContentId Id { get; set; }
            public XDocument Content { get; set;}
            public XmlNamespaceManager Xnm { get; set; }
        }

        private static object _locker = new object();

        private static volatile OneNoteProxy _instance = null;
        public static OneNoteProxy Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (_locker)
                    {
                        if (_instance == null)
                        {
                            _instance = new OneNoteProxy();
                        }
                    }
                }

                return _instance;
            }
        }

        private Dictionary<OneNoteHierarchyContentId, HierarchyElement> _hierarchyContentCache = new Dictionary<OneNoteHierarchyContentId, HierarchyElement>();
        private Dictionary<string, PageContent> _pageContentCache = new Dictionary<string, PageContent>();
        private Dictionary<CommentPageId, string> _commentPagesIds = new Dictionary<CommentPageId, string>();
        private Dictionary<string, OneNoteProxy.BiblePageId> _processedBiblePages = new Dictionary<string, BiblePageId>();

        protected OneNoteProxy()
        {

        }

        public Dictionary<string, OneNoteProxy.BiblePageId> ProcessedBiblePages
        {
            get
            {
                return _processedBiblePages;
            }
        }

        public void AddProcessedBiblePages(string bibleSectionId, string biblePageId, string biblePageName)
        {
            if (!_processedBiblePages.ContainsKey(biblePageId))
            {
                _processedBiblePages.Add(biblePageId, new BiblePageId()
                {
                    SectionId = bibleSectionId,
                    PageId = biblePageId,
                    PageName = biblePageName
                });
            }
        }

        public string GetCommentPageId(Application oneNoteApp, string bibleSectionId, string biblePageId, string biblePageName, string commentPageName)
        {
            CommentPageId key = new CommentPageId()
            {
                BiblePageId = new BiblePageId()
                {
                    SectionId = bibleSectionId,
                    PageId = biblePageId,
                    PageName = biblePageName
                },
                CommentsPageName = commentPageName
            };
            if (!_commentPagesIds.ContainsKey(key))
            {
                string commentPageId = VerseLinkManager.FindVerseLinkPageAndCreateIfNeeded(oneNoteApp, bibleSectionId, biblePageId, biblePageName, commentPageName);
                _commentPagesIds.Add(key, commentPageId);
            }

            return _commentPagesIds[key];
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="oneNoteApp"></param>
        /// <param name="hierarchyId"></param>
        /// <param name="scope"></param>
        /// <param name="refreshCache">Стоит ли загружать данные из OneNote (true) или из кэша (false)</param>
        /// <returns></returns>
        public HierarchyElement GetHierarchy(Application oneNoteApp, string hierarchyId, HierarchyScope scope, bool refreshCache = false)
        {
            OneNoteHierarchyContentId contentId = new OneNoteHierarchyContentId() { ID = hierarchyId, ContentScope = scope };

            HierarchyElement result;

            if (!_hierarchyContentCache.ContainsKey(contentId) || refreshCache)
            {
                lock (_locker)
                {
                    string xml;
                    oneNoteApp.GetHierarchy(hierarchyId, scope, out xml);

                    XmlNamespaceManager xnm;
                    XDocument doc = OneNoteUtils.GetXDocument(xml, out xnm);

                    if (!_hierarchyContentCache.ContainsKey(contentId))
                        _hierarchyContentCache.Add(contentId, new HierarchyElement() { Id = contentId, Content = doc, Xnm = xnm });
                    else
                        _hierarchyContentCache[contentId].Content = doc;
                }
            }
            
            result = _hierarchyContentCache[contentId];
            
            return result;
        }

        public PageContent GetPageContent(Application oneNoteApp, string pageId)
        {
            return GetPageContent(oneNoteApp, pageId, false);
        }


        private PageContent GetPageContent(Application oneNoteApp, string pageId, bool refreshCache)
        {
            PageContent result;

            if (!_pageContentCache.ContainsKey(pageId) || refreshCache)
            {
                lock (_locker)
                {
                    string xml;
                    oneNoteApp.GetPageContent(pageId, out xml);

                    XmlNamespaceManager xnm;
                    XDocument doc = OneNoteUtils.GetXDocument(xml, out xnm);

                    if (!_pageContentCache.ContainsKey(pageId))
                        _pageContentCache.Add(pageId, new PageContent() { PageId = pageId, Content = doc, Xnm = xnm });
                    else
                        _pageContentCache[pageId].Content = doc;
                }
            }

            result = _pageContentCache[pageId];

            return result;
        }

        public void RefreshHierarchyCache(Application oneNoteApp, string hierarchyId, HierarchyScope scope)
        {
            GetHierarchy(oneNoteApp, hierarchyId, scope, true);
        }

        public void CommitAllModifiedPages(Application oneNoteApp, Action<PageContent> onPageProcessed)
        {
            List<string> toRemove = new List<string>();
            foreach (PageContent pageContent in _pageContentCache.Values.Where(pg => pg.WasModified))
            {
                oneNoteApp.UpdatePageContent(pageContent.Content.ToString());
                //pageContent.WasModified = false;
                //GetPageContent(oneNoteApp, pageContent.PageId, true);  // обновляем, так как OneNote сам модифицирует контен страницы при обновлении
                toRemove.Add(pageContent.PageId);

                if (onPageProcessed != null)
                    onPageProcessed(pageContent);
            }

            foreach (var pageId in toRemove)  // нам нет смысла их заново загружать в кэш, так как возможно они больше не понадобятся. 
                _pageContentCache.Remove(pageId);
        }

        //public void RefreshPageContentCache(Application oneNoteApp, string pageId)
        //{
        //    GetPageContent(oneNoteApp, pageId, true);
        //}
    }
}
