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

        public class PageContent
        {            
            public string PageId { get; set; }
            public XDocument Content { get; set; }
            public XmlNamespaceManager Xnm { get; set; }
            public bool WasModified { get; set; }
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

        private Dictionary<OneNoteHierarchyContentId, string> _hierarchyContentCache = new Dictionary<OneNoteHierarchyContentId, string>();
        private Dictionary<string, PageContent> _pageContentCache = new Dictionary<string, PageContent>();

        protected OneNoteProxy()
        {

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="oneNoteApp"></param>
        /// <param name="hierarchyId"></param>
        /// <param name="scope"></param>
        /// <param name="refreshCache">Стоит ли загружать данные из OneNote (true) или из кэша (false)</param>
        /// <returns></returns>
        public string GetHierarchy(Application oneNoteApp, string hierarchyId, HierarchyScope scope, bool refreshCache = false)
        {
            OneNoteHierarchyContentId contentId = new OneNoteHierarchyContentId() { ID = hierarchyId, ContentScope = scope };

            string result;
            if (!_hierarchyContentCache.ContainsKey(contentId) || refreshCache)
            {
                lock (_locker)
                {
                    oneNoteApp.GetHierarchy(hierarchyId, scope, out result);

                    if (!_hierarchyContentCache.ContainsKey(contentId))
                        _hierarchyContentCache.Add(contentId, result);
                    else
                        _hierarchyContentCache[contentId] = result;
                }
            }
            else
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
