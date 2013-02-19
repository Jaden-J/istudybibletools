using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.OneNote;
using System.Xml.Linq;
using BibleCommon.Helpers;
using System.Xml;
using BibleCommon.Common;
using BibleCommon.Consts;
using System.Runtime.InteropServices;
using BibleCommon.Contracts;
using System.Xml.XPath;
using System.IO;
using BibleCommon.Handlers;


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
                int result = this.ContentScope.GetHashCode();

                if (!string.IsNullOrEmpty(this.ID))
                    result = result ^ this.ID.GetHashCode();

                return result;
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
            public bool AddLatestAnalyzeTimeMetaAttribute { get; set; }
        }

        public class HierarchyElement
        {
            public OneNoteHierarchyContentId Id { get; set; }
            public XDocument Content { get; set;}
            public XmlNamespaceManager Xnm { get; set; }
            public bool WasModified { get; set; }
        }

        public class SortPageInfo
        {
            public string SectionId { get; set;}
            public string PageId { get; set;}
            public string ParentPageId { get; set;}
            public int PageLevel { get; set; }
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
        private HashSet<SimpleVersePointer> _processedVerses = new HashSet<SimpleVersePointer>();
        private List<SortPageInfo> _sortVerseLinkPagesInfo = new List<SortPageInfo>();
        private Dictionary<string, string> _bibleVersesLinks = null;
        private Dictionary<string, ModuleDictionaryInfo> _moduleDictionaries = new Dictionary<string, ModuleDictionaryInfo>();
        private Dictionary<string, Dictionary<string, string>> _dictionariesTermsLinks = new Dictionary<string, Dictionary<string, string>>();        

        private bool? _isBibleVersesLinksCacheActive;

        protected OneNoteProxy()
        {

        }

        public DictionaryTermLink GetDictionaryTermLink(string term, string dictionaryModuleShortName)
        {
            Dictionary<string, string> cachedLinks = null;
            if (!_dictionariesTermsLinks.ContainsKey(dictionaryModuleShortName))
            {
                cachedLinks = DictionaryTermsCacheManager.LoadCachedDictionary(dictionaryModuleShortName);
                _dictionariesTermsLinks.Add(dictionaryModuleShortName, cachedLinks);
            }
            else
                cachedLinks = _dictionariesTermsLinks[dictionaryModuleShortName];

            if (!cachedLinks.ContainsKey(term))
                throw new ArgumentException(string.Format(BibleCommon.Resources.Constants.DictionaryTermNotFoundInCache, term, dictionaryModuleShortName));

            return new DictionaryTermLink(cachedLinks[term]);
        }

        public ModuleDictionaryInfo GetModuleDictionary(string moduleShortName)
        {
            ModuleDictionaryInfo result = null;
            if (!_moduleDictionaries.ContainsKey(moduleShortName))
            {
                result = ModulesManager.GetModuleDictionaryInfo(moduleShortName);
                _moduleDictionaries.Add(moduleShortName, result);
            }
            else
                result = _moduleDictionaries[moduleShortName];

            return result;
        }

        public List<SortPageInfo> SortVerseLinkPagesInfo
        {
            get
            {
                return _sortVerseLinkPagesInfo;
            }
        }

        public void RegisterVerseLinkSortPage(string sectionId, string newPageId, string verseLinkParentPageId, int pageLevel)
        {
            _sortVerseLinkPagesInfo.Add(new SortPageInfo()
            {
                SectionId = sectionId,
                PageId = newPageId,
                ParentPageId = verseLinkParentPageId,
                PageLevel = pageLevel
            });
        }     
      
        public string GenerateHref(ref Application oneNoteApp, string pageId, string objectId)
        {
            if (string.IsNullOrEmpty(pageId))
                throw new ArgumentNullException("pageId");

            LinkId key = new LinkId()
            {                
                PageId = pageId,
                ObjectId = objectId
            };

            if (!_linksCache.ContainsKey(key))
            {
                //lock (_locker)
                {
                    string link = null;

                    OneNoteUtils.UseOneNoteAPI(ref oneNoteApp, (oneNoteAppSafe) =>
                    {
                        oneNoteAppSafe.GetHyperlinkToObject(pageId, objectId, out link);
                    });

                    if (SettingsManager.Instance.UseProxyLinksForLinks)
                        link = GetProxyLink(link, pageId, objectId);                    

                    //if (!_linksCache.ContainsKey(key))   // пока в этом нет смысла
                    _linksCache.Add(key, link);
                }
            }

            return _linksCache[key];
        }

        public static bool IsProxyLink(string link)
        {
            return link.IndexOf("&" + Constants.QueryParameterKey_CustomPageId + "=") > -1;
        }

        public static string GetProxyLink(string link, string pageId, string objectId)
        {
            link = string.Concat(
                        link.Replace("onenote:", string.Format("{0}:_", Constants.ISBTOpenProtocol)),
                        "&", Constants.QueryParameterKey_CustomPageId, "=", pageId,
                        "&", Constants.QueryParameterKey_CustomObjectId, "=", objectId);

            return link;
        }

        public Dictionary<string, OneNoteProxy.BiblePageId> BiblePagesWithUpdatedLinksToNotesPages
        {
            get
            {
                return _processedBiblePages;
            }
        }

        public HashSet<SimpleVersePointer> ProcessedVersesOnBiblePagesWithUpdatedLinksToNotesPages
        {
            get
            {
                return _processedVerses;
            }
        }

        public void AddProcessedVerseOnBiblePageWithUpdatedLinksToNotesPages(VersePointer vp, VerseNumber? verseNumber)
        {
            var svp = vp.ToSimpleVersePointer();
            if (verseNumber.HasValue)
                svp.VerseNumber = verseNumber.Value;

            svp.GetAllVerses().ForEach(v =>
            {
                if (!_processedVerses.Contains(v))
                    _processedVerses.Add(v);
            });
        }      

        public void AddProcessedBiblePageWithUpdatedLinksToNotesPages(string bibleSectionId, string biblePageId, string biblePageName, VersePointer chapterPointer)
        {
            if (!_processedBiblePages.ContainsKey(biblePageId))
            {
                //lock (_locker)
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


        public string GetNotesPageId(ref Application oneNoteApp, string bibleSectionId, string biblePageId, string biblePageName,
            string notesPageName, out bool pageWasCreated, string notesParentPageName = null, int pageLevel = 1)
        {
            return GetVerseLinkPageId(ref oneNoteApp, bibleSectionId, biblePageId, biblePageName, notesPageName, true, 
                notesParentPageName, pageLevel, out pageWasCreated, true);
        }

        public string GetCommentPageId(ref Application oneNoteApp, string bibleSectionId, string biblePageId,
            string biblePageName, string commentPageName, out bool pageWasCreated, bool createNewPageIfNeeded = true)
        {
            return GetVerseLinkPageId(ref oneNoteApp, bibleSectionId, biblePageId, biblePageName, commentPageName, false, null, 1, out pageWasCreated, createNewPageIfNeeded);
        }

        private string GetVerseLinkPageId(ref Application oneNoteApp, string bibleSectionId, string biblePageId, string biblePageName, string commentPageName,
            bool isSummaryNotesPage, string verseLinkParentPageName, int pageLevel, out bool pageWasCreated, bool createNewPageIfNeeded)
        {
            pageWasCreated = false;
            string verseLinkParentPageId = null;
            if (!string.IsNullOrEmpty(verseLinkParentPageName))
                verseLinkParentPageId = GetVerseLinkPageId(ref oneNoteApp, bibleSectionId, biblePageId, biblePageName,
                    verseLinkParentPageName, isSummaryNotesPage, null, 1, out pageWasCreated, createNewPageIfNeeded);

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
                //lock (_locker)         // пока в этом нет смысла
                {
                    string commentPageId = VerseLinkManager.FindVerseLinkPageAndCreateIfNeeded(ref oneNoteApp, bibleSectionId, biblePageId, biblePageName,
                        commentPageName, isSummaryNotesPage, out pageWasCreated, verseLinkParentPageId, pageLevel, createNewPageIfNeeded);
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
        public HierarchyElement GetHierarchy(ref Application oneNoteApp, string hierarchyId, HierarchyScope scope, bool refreshCache = false)
        {
            OneNoteHierarchyContentId contentId = new OneNoteHierarchyContentId() { ID = hierarchyId, ContentScope = scope };

            HierarchyElement result;

            if (!_hierarchyContentCache.ContainsKey(contentId) || refreshCache)
            {
                //lock (_locker)
                {
                    string xml = null;
                    try
                    {
                        OneNoteUtils.UseOneNoteAPI(ref oneNoteApp, (oneNoteAppSafe) =>
                        {
                            oneNoteAppSafe.GetHierarchy(hierarchyId, scope, out xml, Constants.CurrentOneNoteSchema);
                        });
                    }
                    catch (Exception ex)
                    {
                        throw new Exception(string.Format(BibleCommon.Resources.Constants.Error_CanNotFindHierarchy, scope, hierarchyId, OneNoteUtils.ParseError(ex.Message)));
                    }

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

        public PageContent GetPageContent(ref Application oneNoteApp, string pageId, PageType pageType, PageInfo pageInfo = PageInfo.piBasic, bool setLineInfo = false)
        {
            return GetPageContent(ref oneNoteApp, pageId, pageType, false, pageInfo, setLineInfo);
        }

        private PageContent GetPageContent(ref Application oneNoteApp, string pageId, PageType pageType, bool refreshCache, PageInfo pageInfo, bool setLineInfo)
        {
            PageContent result;

            var key = string.Format("{0}_{1}", pageId, pageInfo);

            if (!_pageContentCache.ContainsKey(key) || refreshCache)
            {
                //lock (_locker)
                {
                    string xml = null;
                    try
                    {
                        OneNoteUtils.UseOneNoteAPI(ref oneNoteApp, (oneNoteAppSafe) =>
                        {
                            oneNoteAppSafe.GetPageContent(key, out xml, pageInfo, Constants.CurrentOneNoteSchema);
                        });
                    }
                    catch (COMException ex)
                    {
                        if (OneNoteUtils.IsError(ex, Error.hrObjectDoesNotExist))
                            throw new Exception("Page does not exists.");
                        else
                            throw;
                    }

                    XmlNamespaceManager xnm;
                    XDocument doc = OneNoteUtils.GetXDocument(xml, out xnm, setLineInfo);

                    if (!_pageContentCache.ContainsKey(key))
                        _pageContentCache.Add(key, new PageContent() { PageId = key, Content = doc, Xnm = xnm, PageType = pageType });
                    else
                        _pageContentCache[key].Content = doc;
                }
            }

            result = _pageContentCache[key];

            return result;
        }

        public void RefreshHierarchyCache(ref Application oneNoteApp, string hierarchyId, HierarchyScope scope)
        {
            GetHierarchy(ref oneNoteApp, hierarchyId, scope, true);
        }

        public void RefreshHierarchyCache()
        {            
            _hierarchyContentCache.Clear();
        }

        public void CommitModifiedPage(ref Application oneNoteApp, PageContent page, bool throwExceptions)
        {
            if (page == null)
                throw new ArgumentNullException("page");            

            try
            {
                if (page.AddLatestAnalyzeTimeMetaAttribute)
                    OneNoteUtils.UpdateElementMetaData(page.Content.Root, Constants.Key_LatestAnalyzeTime, DateTime.UtcNow.AddSeconds(10).ToString(), page.Xnm);

                OneNoteUtils.UpdatePageContentSafe(ref oneNoteApp, page.Content, page.Xnm);
            }
            catch (Exception ex)
            {
                Logger.LogError(string.Format("{0} '{1}'.", BibleCommon.Resources.Constants.Error_UpdatePage, (string)page.Content.Root.Attribute("name")), ex);
                if (throwExceptions)
                {
                    _pageContentCache.Remove(page.PageId);  // мы всё равно не смогли обновить эту страницу.
                    throw;
                }
            }

            _pageContentCache.Remove(page.PageId);
        }

        public void CommitAllModifiedPages(ref Application oneNoteApp, bool throwExceptions, Func<PageContent, bool> filter, 
            Action<int> onAllPagesToCommitFound, Action<PageContent> onPageProcessed)
        {   
            List<PageContent> toCommit = _pageContentCache.Values.Where(pg => pg.WasModified && (filter == null || filter(pg))).ToList();
            if (onAllPagesToCommitFound != null)
                onAllPagesToCommitFound(toCommit.Count);
            
            foreach (PageContent page in toCommit)
            {
                try
                {
                    if (page.AddLatestAnalyzeTimeMetaAttribute)
                        OneNoteUtils.UpdateElementMetaData(page.Content.Root, Constants.Key_LatestAnalyzeTime, DateTime.UtcNow.AddSeconds(10).ToString(), page.Xnm);

                    OneNoteUtils.UpdatePageContentSafe(ref oneNoteApp, page.Content, page.Xnm);
                }
                catch (Exception ex)
                {
                    Logger.LogError(string.Format("{0} '{1}'.", BibleCommon.Resources.Constants.Error_UpdatePage, (string)page.Content.Root.Attribute("name")), ex);
                    if (throwExceptions)
                    {
                        _pageContentCache.Remove(page.PageId);  // мы всё равно не смогли обновить эту страницу.
                        throw;
                    }
                }

                if (onPageProcessed != null)
                    onPageProcessed(page);
            }

            //lock (_locker)
            {
                foreach (var page in toCommit)  
                    _pageContentCache.Remove(page.PageId);
            }
        }        

        public void CommitAllModifiedHierarchy(ref Application oneNoteApp, Action<int> onAllHierarchyToCommitFound,
            Action<HierarchyElement> onHierarchyElementProcessed)
        {
            var toCommit = _hierarchyContentCache.Values.Where(h => h.WasModified).ToList();

            if (onAllHierarchyToCommitFound != null)
                onAllHierarchyToCommitFound(toCommit.Count);

            foreach (var hierarchy in toCommit)
            {
                OneNoteUtils.UseOneNoteAPI(ref oneNoteApp, (oneNoteAppSafe) =>
                {
                    oneNoteAppSafe.UpdateHierarchy(hierarchy.Content.ToString(), Constants.CurrentOneNoteSchema);
                });

                if (onHierarchyElementProcessed != null)
                    onHierarchyElementProcessed(hierarchy);
            }

            //lock (_locker)
            {
                foreach (var h in toCommit)
                    _hierarchyContentCache.Remove(h.Id);
            }
        }

        //public void RefreshPageContentCache(Application oneNoteApp, string pageId)
        //{
        //    GetPageContent(oneNoteApp, pageId, true);
        //}
        
        public bool IsBibleVersesLinksCacheActive
        {
            get
            {
                if (!_isBibleVersesLinksCacheActive.HasValue)
                    _isBibleVersesLinksCacheActive = BibleVersesLinksCacheManager.CacheIsActive(SettingsManager.Instance.NotebookId_Bible);

                return _isBibleVersesLinksCacheActive.Value;
            }            
        }

        public void CleanBibleVersesLinksCache()
        {
            _bibleVersesLinks = null;
            _isBibleVersesLinksCacheActive = false;
            BibleVersesLinksCacheManager.RemoveCacheFile(SettingsManager.Instance.NotebookId_Bible);            
        }

        public VersePointerLink GetVersePointerLink(VersePointer vp)
        {
            if (_bibleVersesLinks == null)
            {
                try
                {
                    _bibleVersesLinks = BibleVersesLinksCacheManager.LoadBibleVersesLinks(SettingsManager.Instance.NotebookId_Bible);
                }
                catch (Exception ex)
                {
                    BibleCommon.Services.Logger.LogError(ex);
                    throw;
                }                
            }

            var vpString = vp.ToFirstVerseString();
            if (_bibleVersesLinks.ContainsKey(vpString))
                return new VersePointerLink(_bibleVersesLinks[vpString]);
            else
            {
                var chapterPointerString = vp.GetChapterPointer().ToFirstVerseString();
                if (_bibleVersesLinks.ContainsKey(chapterPointerString))
                    return new VersePointerLink(_bibleVersesLinks[chapterPointerString]);
            }

            return null;
        }      
    }
}
