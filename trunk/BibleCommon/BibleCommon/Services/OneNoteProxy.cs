using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.OneNote;

namespace BibleCommon.Services
{
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
        private Dictionary<string, string> _pageContentCache = new Dictionary<string, string>();

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

        public string GetPageContent(Application oneNoteApp, string pageId, bool refreshCache = false)
        {
            string result;
            if (!_pageContentCache.ContainsKey(pageId) || refreshCache)
            {
                lock (_locker)
                {
                    oneNoteApp.GetPageContent(pageId, out result);

                    if (!_pageContentCache.ContainsKey(pageId))
                        _pageContentCache.Add(pageId, result);
                    else
                        _pageContentCache[pageId] = result;
                }
            }
            else
                result = _pageContentCache[pageId];

            return result;
        }

        public void RefreshHierarchyCache(Application oneNoteApp, string hierarchyId, HierarchyScope scope)
        {
            GetHierarchy(oneNoteApp, hierarchyId, scope, true);
        }

        public void RefreshPageContentCache(Application oneNoteApp, string pageId)
        {
            GetPageContent(oneNoteApp, pageId, true);
        }
    }
}
