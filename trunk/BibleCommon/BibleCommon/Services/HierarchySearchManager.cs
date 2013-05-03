using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Microsoft.Office.Interop.OneNote;
using System.Xml;
using System.Xml.XPath;
using BibleCommon.Common;
using BibleCommon.Helpers;
using System.IO;


namespace BibleCommon.Services
{
    public static class HierarchySearchManager
    {   
        public enum FindVerseLevel
        {
            OnlyFirstVerse,
            OnlyVersesOfFirstChapter,    // пока не работает, если указана ссылка, включающая в себя несколько глав (например, 5:6-6:7)
            AllVerses
        }
        
        [Serializable]
        public class HierarchySearchResult : BibleSearchResult
        {
            public HierarchySearchResult()
            {
                HierarchyObjectInfo = new BibleHierarchyObjectInfo();
            }

            /// <summary>
            /// 
            /// </summary>
            /// <param name="oneNoteApp"></param>
            /// <param name="vp"></param>
            /// <param name="versePointerLink"></param>
            /// <param name="findAllVerseObjects">Поиск осуществляется только по текущей главе. Чтобы найти все стихи во всех главах (если например ссылка 3:4-6:8), то надо отдельно вызвать GetAllIncludedVersesExceptFirst</param>
            public HierarchySearchResult(ref Application oneNoteApp, string bibleNotebookId, VersePointer vp, VersePointerLink versePointerLink, FindVerseLevel findAllVerseObjects, bool loadedFromCache)
                : this()
            {   
                this.ResultType = vp.IsChapter == versePointerLink.IsChapter ? BibleHierarchySearchResultType.Successfully : BibleHierarchySearchResultType.PartlyFound;
                this.HierarchyStage = versePointerLink.IsChapter ? BibleHierarchyStage.Page : BibleHierarchyStage.ContentPlaceholder;
                this.HierarchyObjectInfo = new BibleHierarchyObjectInfo()
                {                    
                    SectionId = versePointerLink.SectionId,
                    PageId = versePointerLink.PageId,
                    PageName = versePointerLink.PageName,
                    VerseInfo = new VerseObjectInfo(versePointerLink),
                    LoadedFromCache = loadedFromCache       
                };

                if (vp.IsMultiVerse &&
                                (findAllVerseObjects == FindVerseLevel.AllVerses || findAllVerseObjects == FindVerseLevel.OnlyVersesOfFirstChapter))
                {
                    List<VersePointer> verses;
                    if (findAllVerseObjects == FindVerseLevel.OnlyVersesOfFirstChapter)
                    {
                        Application temp = null;
                        verses = vp.GetAllIncludedVersesExceptFirst(ref temp,
                                                        new GetAllIncludedVersesExceptFirstArgs() { SearchOnlyForFirstChapter = true });
                    }
                    else
                        verses = vp.GetAllIncludedVersesExceptFirst(ref oneNoteApp,
                                                        new GetAllIncludedVersesExceptFirstArgs() { BibleNotebookId = bibleNotebookId });

                    foreach (var subVp in verses)
                    {
                        var subLink = ApplicationCache.Instance.GetVersePointerLink(subVp);
                        if (subLink != null)
                            this.HierarchyObjectInfo.AdditionalObjectsIds.Add(subVp, new VerseObjectInfo(subLink));
                    }
                }
            }
        }       

        /// <summary>
        /// 
        /// </summary>
        /// <param name="oneNoteApp"></param>
        /// <param name="bibleNotebookId"></param>
        /// <param name="vp"></param>
        /// <param name="findAllVerseObjects">Поиск осуществляется только по текущей главе. Чтобы найти все стихи во всех главах (если например ссылка 3:4-6:8), то надо отдельно вызвать GetAllIncludedVersesExceptFirst</param>
        /// <returns></returns>
        public static HierarchySearchResult GetHierarchyObject(ref Application oneNoteApp, string bibleNotebookId, ref VersePointer vp, FindVerseLevel findAllVerseObjects,
            string pageId, string contentObjectId, bool useCacheIfAvailable = true)
        {
            var result = GetHierarchyObjectInternal(ref oneNoteApp, bibleNotebookId, vp, findAllVerseObjects, useCacheIfAvailable);

            if (!result.FoundSuccessfully)
            {
                BibleCommon.Services.Logger.LogWarning(pageId, contentObjectId, BibleCommon.Resources.Constants.VerseNotFound, vp.OriginalVerseName);
            }
            else
            {
                //ответ на вопрос "почему для разных ситуаций (есть кэш/нет кэша; Иуд 1-3/Иуд 2-3) используется разный подход?". Потому что, если в кэше ничего не нашли (Иуд 2-3), то начинаем искать в структуре (если нет кэша), а это долго. Если же в кэше нашли, а дополнительные ссылки не нашли (Иуд 1-3), то в структуре уже ничего не ищем (если нет кэша), потому здесь можно использовать общий подход (который и приведён ниже)
                if (vp.IsChapter && (vp.ParentVersePointer ?? vp).TopChapter.HasValue 
                    && (findAllVerseObjects == FindVerseLevel.AllVerses || findAllVerseObjects == FindVerseLevel.OnlyVersesOfFirstChapter)
                    && result.HierarchyObjectInfo.AdditionalObjectsIds.Count == 0)
                {                                                                                                         // возможно имеем дело со стихами "Иуд 1-3"
                    if (BookHasOnlyOneChapter(ref oneNoteApp, vp, result.HierarchyObjectInfo, useCacheIfAvailable))                    
                    {
                        var changedResult = TryToChangeVerseAsOneChapteredBookAndSearchInHierarchy(
                                               ref oneNoteApp, bibleNotebookId, ref vp, findAllVerseObjects, useCacheIfAvailable, true);
                        if (changedResult != null)
                            result = changedResult;
                    }
                }
            }

            return result;
        }

        private static HierarchySearchResult GetHierarchyObjectInternal(ref Application oneNoteApp, string bibleNotebookId, VersePointer vp, 
            FindVerseLevel findAllVerseObjects, bool useCacheIfAvailable = true)
        {   
            var result = new HierarchySearchResult();
            result.ResultType = BibleHierarchySearchResultType.NotFound;

            if (!vp.IsValid)
                throw new ArgumentException("versePointer is not valid");

            if (useCacheIfAvailable && ApplicationCache.Instance.IsBibleVersesLinksCacheActive)
            {                
                var versePointerLink = ApplicationCache.Instance.GetVersePointerLink(vp);
                if (versePointerLink != null)
                    return new HierarchySearchResult(ref oneNoteApp, bibleNotebookId, vp, versePointerLink, findAllVerseObjects, true);
                else if (vp.IsChapter) // возможно стих типа "2Ин 8"
                {
                    if (BookHasOnlyOneChapter(ref oneNoteApp, vp, result.HierarchyObjectInfo, useCacheIfAvailable))  // result.HierarchyObjectInfo - пустое, но оно там и не нужно.
                    {
                        var changedVerseResult = TryToChangeVerseAsOneChapteredBookAndSearchInHierarchy(
                            ref oneNoteApp, bibleNotebookId, ref vp, findAllVerseObjects, useCacheIfAvailable, false);
                        if (changedVerseResult != null)
                            return changedVerseResult;
                    }
                }
            }
            else
                result.HierarchyObjectInfo.LoadedFromCache = false;


            XElement targetSection = FindBibleBookSection(ref oneNoteApp, bibleNotebookId, vp.Book.SectionName);
            if (targetSection != null)
            {
                result.HierarchyObjectInfo.SectionId = (string)targetSection.Attribute("ID");
                result.ResultType = BibleHierarchySearchResultType.Successfully;
                result.HierarchyStage = BibleHierarchyStage.Section;                

                XElement targetPage = HierarchySearchManager.FindPage(ref oneNoteApp, result.HierarchyObjectInfo.SectionId, vp.Chapter.Value);

                if (targetPage == null && vp.IsChapter)  // возможно стих типа "2Ин 8"
                {
                    if (BookHasOnlyOneChapter(ref oneNoteApp, vp, result.HierarchyObjectInfo, useCacheIfAvailable))                    
                    {
                        var changedVerseResult = TryToChangeVerseAsOneChapteredBookAndSearchInHierarchy(
                            ref oneNoteApp, bibleNotebookId, ref vp, findAllVerseObjects, useCacheIfAvailable, false);
                        if (changedVerseResult != null)
                            return changedVerseResult;                     
                    }
                }

                if (targetPage != null)
                {                    
                    result.HierarchyObjectInfo.PageId = (string)targetPage.Attribute("ID");
                    result.HierarchyObjectInfo.PageName = (string)targetPage.Attribute("name");
                    result.HierarchyStage = BibleHierarchyStage.Page;

                    var pageContent = ApplicationCache.Instance.GetPageContent(ref oneNoteApp, result.HierarchyObjectInfo.PageId, ApplicationCache.PageType.Bible);
                    VerseNumber? verseNumber;
                    string verseTextWithoutNumber;
                    XElement targetVerseEl = FindVerse(pageContent.Content, vp.IsChapter, vp.Verse.Value, pageContent.Xnm, out verseNumber, out verseTextWithoutNumber);

                    if (targetVerseEl != null)
                    {
                        result.HierarchyObjectInfo.VerseInfo = new VerseObjectInfo()
                                                                        {
                                                                            ObjectId = (string)targetVerseEl.Parent.Attribute("objectID"),
                                                                            VerseNumber = verseNumber
                                                                        };
                        if (!vp.IsChapter)
                        {
                            result.HierarchyStage = BibleHierarchyStage.ContentPlaceholder;

                            if (vp.IsMultiVerse &&
                                (findAllVerseObjects == FindVerseLevel.AllVerses || findAllVerseObjects == FindVerseLevel.OnlyVersesOfFirstChapter))
                            {
                                LoadAdditionalVersesInfo(ref oneNoteApp, bibleNotebookId, vp, findAllVerseObjects, pageContent, ref result);                              
                            }
                        }
                    }
                    else if (!vp.IsChapter)   // Если по идее должны были найти, а не нашли...
                    {
                        result.ResultType = BibleHierarchySearchResultType.PartlyFound;
                    }                    
                }
            }

            return result;
        }

        private static bool BookHasOnlyOneChapter(ref Application oneNoteApp, VersePointer vp, BibleHierarchyObjectInfo hierarchyObjectInfo, bool useCacheIfAvailable)
        {
            if (useCacheIfAvailable && ApplicationCache.Instance.IsBibleVersesLinksCacheActive)
                return ApplicationCache.Instance.GetVersePointerLink(new VersePointer(vp.Book.Name, 2)) == null;
            else            
                return GetBookChaptersCount(ref oneNoteApp, hierarchyObjectInfo.SectionId) == 1;                                               
        }

        private static HierarchySearchResult TryToChangeVerseAsOneChapteredBookAndSearchInHierarchy(ref Application oneNoteApp, string bibleNotebookId,
                        ref VersePointer vp, FindVerseLevel findAllVerseObjects, bool useCacheIfAvailable, bool checkAdditionalVersesCount)
        {
            var modifiedVp = new VersePointer(vp.OriginalVerseName);
            modifiedVp.ChangeVerseAsOneChapteredBook();
            var changedVerseResult = GetHierarchyObjectInternal(ref oneNoteApp, bibleNotebookId, modifiedVp, findAllVerseObjects, useCacheIfAvailable);
            if (changedVerseResult.FoundSuccessfully 
                && (!checkAdditionalVersesCount || changedVerseResult.HierarchyObjectInfo.AdditionalObjectsIds.Count > 0))
            {
                vp.ChangeVerseAsOneChapteredBook();
                return changedVerseResult;
            }
            else
                return null;
        }

        private static void LoadAdditionalVersesInfo(ref Application oneNoteApp, string bibleNotebookId, VersePointer vp, FindVerseLevel findAllVerseObjects,
            ApplicationCache.PageContent chapterPageContent, ref HierarchySearchResult result)
        {
            List<VersePointer> verses;
            if (findAllVerseObjects == FindVerseLevel.OnlyVersesOfFirstChapter)
            {
                Application temp = null;
                verses = vp.GetAllIncludedVersesExceptFirst(ref temp,
                                                new GetAllIncludedVersesExceptFirstArgs() { SearchOnlyForFirstChapter = true });
            }
            else
                verses = vp.GetAllIncludedVersesExceptFirst(ref oneNoteApp,
                                                new GetAllIncludedVersesExceptFirstArgs() { BibleNotebookId = bibleNotebookId });

            foreach (var additionalVerse in verses)
            {
                VerseNumber? additionalVerseNumber;
                string additionalVerseTextWithoutNumber;
                var additionalPageContent = chapterPageContent;
                if (additionalVerse.Chapter != vp.Chapter)
                {
                    var additionalPageEl = HierarchySearchManager.FindPage(ref oneNoteApp, result.HierarchyObjectInfo.SectionId, additionalVerse.Chapter.Value);
                    if (additionalPageEl == null)
                        continue;
                    var additionalPageId = (string)additionalPageEl.Attribute("ID");
                    additionalPageContent = ApplicationCache.Instance.GetPageContent(ref oneNoteApp, additionalPageId, ApplicationCache.PageType.Bible);
                }

                var additionalVeseEl = FindVerse(additionalPageContent.Content, vp.IsChapter, additionalVerse.Verse.Value, additionalPageContent.Xnm,
                    out additionalVerseNumber, out additionalVerseTextWithoutNumber);
                if (additionalVeseEl != null)
                {
                    result.HierarchyObjectInfo.AdditionalObjectsIds.Add(additionalVerse,
                                                                        new VerseObjectInfo()
                                                                        {
                                                                            ObjectId = (string)additionalVeseEl.Parent.Attribute("objectID"),
                                                                            VerseNumber = additionalVerseNumber
                                                                        });
                }
                else
                    continue;
            }
        }

        internal static XElement FindChapterPage(ref Application oneNoteApp, string bibleNotebookId, string sectionName, int chapterIndex)
        {
            XElement targetSection = FindBibleBookSection(ref oneNoteApp, bibleNotebookId, sectionName);
            if (targetSection != null)
            {
                return HierarchySearchManager.FindPage(ref oneNoteApp, (string)targetSection.Attribute("ID"), chapterIndex);                
            }

            return null;
        }        

        internal static XElement FindVerse(XDocument pageContent, bool isChapter, int verse, XmlNamespaceManager xnm, out VerseNumber? verseNumber, out string verseTextWithoutNumber)
        {
            XElement pointerElement = null;
            verseNumber = null;
            verseTextWithoutNumber = null;

            if (!isChapter)
            {
                string[] searchPatterns = new string[] { 
                    "//one:Outline/one:OEChildren/one:OE/one:Table/one:Row/one:Cell[1]/one:OEChildren/one:OE/one:T[starts-with(.,'{0} ')]",
                    "//one:Outline/one:OEChildren/one:OE/one:T[starts-with(.,'{0} ')]",
                    "//one:Outline/one:OEChildren/one:OE/one:Table/one:Row/one:Cell[1]/one:OEChildren/one:OE/one:T[starts-with(.,'{0}&nbsp;')]",
                    "//one:Outline/one:OEChildren/one:OE/one:T[starts-with(.,'{0}&nbsp;')]",
                    "//one:Outline/one:OEChildren/one:OE/one:Table/one:Row/one:Cell[1]/one:OEChildren/one:OE/one:T[contains(.,'>{0}<')]",
                    "//one:Outline/one:OEChildren/one:OE/one:T[contains(.,'>{0}<')]",
                    "//one:Outline/one:OEChildren/one:OE/one:Table/one:Row/one:Cell[1]/one:OEChildren/one:OE/one:T[contains(.,'>{0}&nbsp;')]",
                    "//one:Outline/one:OEChildren/one:OE/one:T[contains(.,'>{0}&nbsp;')]",
                    "//one:Outline/one:OEChildren/one:OE/one:Table/one:Row/one:Cell[1]/one:OEChildren/one:OE/one:T[contains(.,'>{0} ')]",
                    "//one:Outline/one:OEChildren/one:OE/one:T[contains(.,'>{0} ')]" };

                foreach (string pattern in searchPatterns)
                {
                    pointerElement = pageContent.Root.XPathSelectElement(string.Format(pattern, verse), xnm);

                    if (pointerElement != null)
                    {                        
                        break;
                    }
                }

                verseNumber = null;
                if (pointerElement != null)                
                    verseNumber = VerseNumber.GetFromVerseText(pointerElement.Value, out verseTextWithoutNumber);                

                if (pointerElement == null || verseNumber == null || !verseNumber.Value.IsVerseBelongs(verse))
                    pointerElement = FindVerseWithIterate(pageContent, verse, xnm, out verseNumber, out verseTextWithoutNumber);                
            }
            else               // тогда возвращаем хотя бы ссылку на заголовок
            {
                pointerElement = NotebookGenerator.GetPageTitle(pageContent, xnm);
            }

            return pointerElement;
        }

        private static XElement FindVerseWithIterate(XDocument pageContent, int verse, XmlNamespaceManager xnm, 
            out VerseNumber? verseNumber, out string verseTextWithoutNumber)
        {
            verseNumber = null;
            verseTextWithoutNumber = null;

            foreach (var textEl in pageContent.Root.XPathSelectElements("//one:Outline/one:OEChildren/one:OE/one:Table/one:Row/one:Cell[1]/one:OEChildren/one:OE/one:T", xnm)
                .Union(pageContent.Root.XPathSelectElements("//one:Outline/one:OEChildren/one:OE/one:T", xnm))) 
            {
                verseNumber = VerseNumber.GetFromVerseText(textEl.Value, out verseTextWithoutNumber);
                if (verseNumber.HasValue && verseNumber.Value.IsVerseBelongs(verse))
                    return textEl;
            }

            return null;
        }

        private static XElement FindPage(ref Application oneNoteApp, string sectionId, int chapter)
        {
            ApplicationCache.HierarchyElement sectionDocument = ApplicationCache.Instance.GetHierarchy(ref oneNoteApp, sectionId, HierarchyScope.hsPages);

            return FindChapterPage(sectionDocument.Content.Root, chapter, sectionDocument.Xnm);
        }

        internal static XElement FindChapterPage(XElement sectionPagesEl, int chapter, XmlNamespaceManager xnm)
        {
            XElement page = sectionPagesEl.XPathSelectElement(string.Format("one:Page[starts-with(@name,'{0} ') and @pageLevel=2]", chapter), xnm);  // "@pageLevel=2" - это условие нам надо, чтобы найти главу, например, "2 глава. 2 Петра", а не просто "2 Петра", 
            if (page == null)
                page = sectionPagesEl.XPathSelectElement(string.Format("one:Page[starts-with(@name,'{0} ')]", chapter), xnm);

            if (page == null)  // нужно для Псалтыря, потому что там главы называются, например "Псалом 5"
                page = sectionPagesEl.XPathSelectElement(
                    string.Format("one:Page[' {0}'=substring(@name,string-length(@name)-{1})]", chapter, chapter.ToString().Length), xnm);

            return page;
        }

        internal static int GetBookChaptersCount(ref Application oneNoteApp, string sectionId)
        {
            if (string.IsNullOrEmpty(sectionId))
                throw new ArgumentNullException("sectionId");

            var sectionDocument = ApplicationCache.Instance.GetHierarchy(ref oneNoteApp, sectionId, HierarchyScope.hsPages);

            return Convert.ToInt32(sectionDocument.Content.Root.XPathEvaluate("count(one:Page)-1", sectionDocument.Xnm));
        }

        public static XElement FindBibleBookSection(ref Application oneNoteApp, string notebookId, string bookSectionName)
        {
            ApplicationCache.HierarchyElement document = ApplicationCache.Instance.GetHierarchy(ref oneNoteApp, notebookId, HierarchyScope.hsSections);

            var targetSection = document.Content.Root.XPathSelectElement(
                string.Format("{0}one:SectionGroup[{2}]/one:Section[@name=\"{1}\"]",
                    !string.IsNullOrEmpty(SettingsManager.Instance.SectionGroupId_Bible)
                        ? string.Format("one:SectionGroup[@ID=\"{0}\"]/", SettingsManager.Instance.SectionGroupId_Bible)
                        : string.Empty,
                    bookSectionName,
                    OneNoteUtils.NotInRecycleXPathCondition),
                document.Xnm);

            if (targetSection == null)  // возможно в другом регистре стало название раздела
            {
                var functionsManager = new CustomXPathFunctions(document.Xnm);
                XmlDocument xd = new XmlDocument();
                xd.Load(document.Content.CreateReader());

                var node = xd.FirstChild.SelectSingleNode(string.Format("{0}one:SectionGroup[{2}]/one:Section[equals(@name, \"{1}\")]",
                        !string.IsNullOrEmpty(SettingsManager.Instance.SectionGroupId_Bible)
                            ? string.Format("one:SectionGroup[@ID=\"{0}\"]/", SettingsManager.Instance.SectionGroupId_Bible)
                            : string.Empty,
                        bookSectionName,
                        OneNoteUtils.NotInRecycleXPathCondition),
                    functionsManager);

                if (node != null)
                {
                    var sectionId = node.Attributes["ID"].Value;
                    targetSection = document.Content.Root.XPathSelectElement(string.Format("//one:Section[@ID=\"{0}\"]", sectionId), document.Xnm);
                }
            }            

            return targetSection;
        }

        public static int? GetChapterVersesCount(ref Application oneNoteApp, string bibleNotebookId, VersePointer versePointer, string pageId, string contentObjectId)
        {
            int? result = null;

            if (bibleNotebookId == SettingsManager.Instance.NotebookId_Bible 
                && SettingsManager.Instance.CanUseBibleContent)
            {
                result = ModulesManager.GetChapterVersesCount(SettingsManager.Instance.CurrentBibleContentCached, versePointer);
            }

            if (result == null)
            {
                var chapterPageResult = GetHierarchyObject(ref oneNoteApp, bibleNotebookId, ref versePointer, FindVerseLevel.OnlyFirstVerse, pageId, contentObjectId);
                if (chapterPageResult.ResultType != BibleHierarchySearchResultType.NotFound 
                    && chapterPageResult.HierarchyStage == BibleHierarchyStage.Page)
                {
                    var pageContent = ApplicationCache.Instance.GetPageContent(ref oneNoteApp, chapterPageResult.HierarchyObjectInfo.PageId, ApplicationCache.PageType.Bible);
                    var table = NotebookGenerator.GetPageTable(pageContent.Content, pageContent.Xnm);
                    if (table != null)                    
                        return table.XPathSelectElements("one:Row", pageContent.Xnm).Count();                    
                }
            }

            return result;
        }

        /// <summary>
        /// Если кэш устареет, то мы его удалим и достанем инфу не из кэша
        /// </summary>
        /// <param name="oneNoteApp"></param>
        /// <param name="bibleNotebookId"></param>
        /// <param name="vp"></param>
        /// <param name="action"></param>
        public static void UseHierarchyObjectSafe(ref Application oneNoteApp, ref BibleHierarchyObjectInfo verseHierarchyObjectInfo, ref VersePointer vp, Func<BibleHierarchyObjectInfo, bool> action,
            string pageId, string contentObjectId)
        {
            try
            {
                action(verseHierarchyObjectInfo);
            }
            catch (NotFoundPageException)
            {
                // возможно дело в устаревшем кэше       

                if (verseHierarchyObjectInfo.LoadedFromCache)
                {
                    var fullSearchResult = HierarchySearchManager.GetHierarchyObject(
                                                ref oneNoteApp, SettingsManager.Instance.NotebookId_Bible, ref vp, HierarchySearchManager.FindVerseLevel.AllVerses, pageId, contentObjectId, false);

                    if (fullSearchResult.FoundSuccessfully)
                    {
                        if (fullSearchResult.HierarchyObjectInfo.PageId != verseHierarchyObjectInfo.PageId)  // если нашли другой ID     // здесь раньше было: if (fullSearchResult.HierarchyObjectInfo.PageId != notePageId.Id)
                        {                            
                            if (action(fullSearchResult.HierarchyObjectInfo))   // значит действительно кэш устарел. Надо его удалить и написать об этом предупреждение                                                   
                            {
                                verseHierarchyObjectInfo = fullSearchResult.HierarchyObjectInfo;
                                ApplicationCache.Instance.CleanBibleVersesLinksCache(false);
                                Logger.LogWarning(BibleCommon.Resources.Constants.BibleVersesLinksCacheWasCleaned);
                            }
                        }
                    }
                }
            }
        }
    }
}
