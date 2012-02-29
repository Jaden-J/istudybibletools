using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.OneNote;
using System.Xml.Linq;
using BibleCommon.Helpers;
using System.Xml;
using BibleCommon.Common;

namespace BibleCommon.Services
{
    /// <summary>
    /// Кэш OneNote
    /// </summary>
    public class OneNoteProxy
    {

        #region Helper classes

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
            public VersePointer ChapterPointer { get; set; }

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

        public enum PageType
        {
            Bible,
            NotePage,
            NotesPage,
            CommentPage
        }

        public class PageContent
        {            
            public string PageId { get; set; }
            public XDocument Content { get; set; }
            public XmlNamespaceManager Xnm { get; set; }
            public bool WasModified { get; set; }
            public PageType PageType { get; set; }
        }

        public class HierarchyElement
        {
            public OneNoteHierarchyContentId Id { get; set; }
            public XDocument Content { get; set;}
            public XmlNamespaceManager Xnm { get; set; }
        }


        public class LinkId
        {   
            public string PageId { get; set; }
            public string ObjectId { get; set; }

            public override int GetHashCode()
            {
                int result = this.PageId.GetHashCode();

                if (!string.IsNullOrEmpty(this.ObjectId))
                    result = result ^ this.ObjectId.GetHashCode();

                return result;
            }

            public override bool Equals(object obj)
            {
                LinkId otherObj = (LinkId)obj;

                return this.PageId == otherObj.PageId
                    && this.ObjectId == otherObj.ObjectId;
            }
        }

        #endregion

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
        private Dictionary<CommentPageId, string> _commentPagesIdsCache = new Dictionary<CommentPageId, string>();
        private Dictionary<string, OneNoteProxy.BiblePageId> _processedBiblePages = new Dictionary<string, BiblePageId>();
        private Dictionary<LinkId, string> _linksCache = new Dictionary<LinkId, string>();
        private Dictionary<string, HashSet<VersePointer>> _processedVerses = new Dictionary<string, HashSet<VersePointer>>();

        protected OneNoteProxy()
        {

        }

        public string GenerateHref(Application oneNoteApp, string pageId, string objectId)
        {
            LinkId key = new LinkId()
            {                
                PageId = pageId,
                ObjectId = objectId
            };

            if (!_linksCache.ContainsKey(key))
            {
                lock (_locker)
                {
                    string link;
                    oneNoteApp.GetHyperlinkToObject(pageId, objectId, out link);
                    
                    //if (!_linksCache.ContainsKey(key))   // пока в этом нет смысла
                        _linksCache.Add(key, link);
                }
            }

            return _linksCache[key];
        }

        public Dictionary<string, OneNoteProxy.BiblePageId> ProcessedBiblePages
        {
            get
            {
                return _processedBiblePages;
            }
        }

        public Dictionary<string, HashSet<VersePointer>> ProcessedVerses
        {
            get
            {
                return _processedVerses;
            }
        }

        public void AddProcessedVerse(string notesPageName, VersePointer vp)
        {
            if (!_processedVerses.ContainsKey(notesPageName))
            {
                lock (_locker)
                {
                    _processedVerses.Add(notesPageName, new HashSet<VersePointer>());
                }
            }

            if (!_processedVerses[notesPageName].Contains(vp))   // отслеживаем обработанные стихи для каждой из страниц сводной заметок
            {
                lock (_locker)
                {
                    _processedVerses[notesPageName].Add(vp);
                }
            }
        }

        public bool ContainsProcessedVerse(string notesPageName, VersePointer vp)
        {
            if (!_processedVerses.ContainsKey(notesPageName))
            {
                lock (_locker)
                {
                    _processedVerses.Add(notesPageName, new HashSet<VersePointer>());
                }
            }

            return _processedVerses[notesPageName].Contains(vp);
        }

        public void AddProcessedBiblePages(string bibleSectionId, string biblePageId, string biblePageName, VersePointer chapterPointer)
        {
            if (!_processedBiblePages.ContainsKey(biblePageId))
            {
                lock (_locker)
                {
                    //if (!_processedBiblePages.ContainsKey(biblePageId))  // пока в этом нет смысла
                    {
                        _processedBiblePages.Add(biblePageId, new BiblePageId()
                        {
                            SectionId = bibleSectionId,
                            PageId = biblePageId,
                            PageName = biblePageName,
                            ChapterPointer = chapterPointer
                        });
                    }
                }
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
            if (!_commentPagesIdsCache.ContainsKey(key))
            {
                lock (_locker)
                {
                    string commentPageId = VerseLinkManager.FindVerseLinkPageAndCreateIfNeeded(oneNoteApp, bibleSectionId, biblePageId, biblePageName, commentPageName);
                    //if (!_commentPagesIdsCache.ContainsKey(key))     // пока в этом нет смысла
                        _commentPagesIdsCache.Add(key, commentPageId);
                }
            }

            return _commentPagesIdsCache[key];
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

        public PageContent GetPageContent(Application oneNoteApp, string pageId, PageType pageType)
        {
            return GetPageContent(oneNoteApp, pageId, pageType, false);
        }

        private PageContent GetPageContent(Application oneNoteApp, string pageId, PageType pageType, bool refreshCache)
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
                        _pageContentCache.Add(pageId, new PageContent() { PageId = pageId, Content = doc, Xnm = xnm, PageType = pageType });
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

        public void CommitAllModifiedPages(Application oneNoteApp, Func<PageContent, bool> filter, 
            Action<int> onAllPagesToCommitFound, Action<PageContent> onPageProcessed)
        {   
            List<PageContent> toCommit = _pageContentCache.Values.Where(pg => pg.WasModified && (filter == null || filter(pg))).ToList();
            if (onAllPagesToCommitFound != null)
                onAllPagesToCommitFound(toCommit.Count);
            
            foreach (PageContent page in toCommit)
            {
                oneNoteApp.UpdatePageContent(page.Content.ToString());

                if (onPageProcessed != null)
                    onPageProcessed(page);
            }

            lock (_locker)
            {
                foreach (var page in toCommit)  
                    _pageContentCache.Remove(page.PageId);
            }
        }

        //public void RefreshPageContentCache(Application oneNoteApp, string pageId)
        //{
        //    GetPageContent(oneNoteApp, pageId, true);
        //}
    }
}
