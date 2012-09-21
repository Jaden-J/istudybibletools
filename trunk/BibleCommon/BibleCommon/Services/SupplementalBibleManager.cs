using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.OneNote;
using System.Xml.Linq;
using BibleCommon.Helpers;
using System.IO;
using BibleCommon.Consts;
using System.Xml;
using BibleCommon.Common;
using System.Xml.XPath;
using BibleCommon.Contracts;

namespace BibleCommon.Services
{
    public static class SupplementalBibleManager
    {
        public static void CreateSupplementalBible(Application oneNoteApp, string moduleShortName, string notebookDirectory, ICustomLogger logger)
        {
            if (string.IsNullOrEmpty(SettingsManager.Instance.GetValidSupplementalBibleNotebookId(oneNoteApp, true)))
            {
                SettingsManager.Instance.NotebookId_SupplementalBible
                    = NotebookGenerator.CreateNotebook(oneNoteApp, Resources.Constants.SupplementalBibleName, notebookDirectory);                
            }
            else
                throw new InvalidOperationException("Supplemental Bible already exists");

            string currentSectionGroupId = null;
            var moduleInfo = ModulesManager.GetModuleInfo(moduleShortName);
            var bibleInfo = ModulesManager.GetModuleBibleInfo(moduleShortName);

            for (int i = 0; i < moduleInfo.BibleStructure.BibleBooks.Count; i++)
            {
                var bibleBookInfo = moduleInfo.BibleStructure.BibleBooks[i];

                bibleBookInfo.SectionName = NotebookGenerator.GetBibleBookSectionName(bibleBookInfo.Name, i, moduleInfo.BibleStructure.OldTestamentBooksCount);

                currentSectionGroupId = GetCurrentSectionGroupId(oneNoteApp, currentSectionGroupId, moduleInfo, i);

                var bookSectionId = NotebookGenerator.AddSection(oneNoteApp, currentSectionGroupId, bibleBookInfo.SectionName);

                var bibleBook = bibleInfo.Content.Books.FirstOrDefault(book => book.Index == bibleBookInfo.Index);
                if (bibleBook == null)
                    throw new Exception("Manifest.xml has Bible books that do not exist in bible.xml");

                foreach (var chapter in bibleBook.Chapters)
                {
                    if (logger != null)
                        logger.LogMessage("{0} '{1} {2}'", BibleCommon.Resources.Constants.ProcessChapter, bibleBookInfo.Name, chapter.Index);

                    GenerateChapterPage(oneNoteApp, chapter, bookSectionId, moduleInfo, bibleBookInfo, bibleInfo);
                }
            }

            oneNoteApp.SyncHierarchy(SettingsManager.Instance.NotebookId_SupplementalBible);

            SettingsManager.Instance.SupplementalBibleModules.Clear();
            SettingsManager.Instance.SupplementalBibleModules.Add(moduleShortName);            
            SettingsManager.Instance.Save();

            if (moduleInfo.Type == ModuleType.Strong)
            {
                DictionaryManager.AddDictionary(oneNoteApp, moduleShortName, notebookDirectory);
            }
        }

        public static BibleParallelTranslationConnectionResult LinkSupplementalBibleWithMainBible(Application oneNoteApp, int supplementalModuleIndex, 
            Dictionary<string, string> strongTermLinksCache, ICustomLogger logger)
        {
            if (supplementalModuleIndex != 0)
                throw new NotSupportedException("supplementalModuleIndex != 0");            

            if (string.IsNullOrEmpty(SettingsManager.Instance.GetValidSupplementalBibleNotebookId(oneNoteApp, true)) 
                || SettingsManager.Instance.SupplementalBibleModules.Count == 0)
                throw new NotConfiguredException("Supplemental Bible does not exists.");

            XmlNamespaceManager xnm = OneNoteUtils.GetOneNoteXNM();
            var nms = XNamespace.Get(Constants.OneNoteXmlNs);

            string supplementalModuleShortName = SettingsManager.Instance.SupplementalBibleModules[supplementalModuleIndex];            

            BibleParallelTranslationManager.MergeModuleWithMainBible(supplementalModuleShortName);                       
            OneNoteLocker.UnlockAllBible(oneNoteApp);            

            BibleParallelTranslationConnectionResult result;
            using (var bibleTranslationManager = new BibleParallelTranslationManager(oneNoteApp,
                            supplementalModuleShortName, SettingsManager.Instance.ModuleName,
                            SettingsManager.Instance.NotebookId_SupplementalBible))
            {
                if (bibleTranslationManager.BaseModuleInfo.Type == ModuleType.Strong)
                    if (strongTermLinksCache == null)
                        throw new ArgumentNullException("strongTermLinksCache");                

                bibleTranslationManager.Logger = logger;                               

                var linkResult = new List<Exception>();

                result = bibleTranslationManager.IterateBaseBible(
                    chapterPageDoc =>
                    {
                        OneNoteProxy.Instance.CommitAllModifiedPages(oneNoteApp, pageContent => pageContent.PageType == OneNoteProxy.PageType.Bible, null, null);

                        return new BibleIteratorArgs() { ChapterDocument = chapterPageDoc };
                    }, true, true,
                    (baseVersePointer, parallelVerse, bibleIteratorArgs) =>
                    {                        
                        linkResult.AddRange(LinkdMainBibleAndSupplementalVerses(oneNoteApp, baseVersePointer, parallelVerse, bibleIteratorArgs, 
                                        bibleTranslationManager.BaseModuleInfo.Type == ModuleType.Strong, strongTermLinksCache, bibleTranslationManager.BaseModuleInfo.BibleStructure.Alphabet, xnm, nms));                       
                    });

                result.Errors.AddRange(linkResult);
            }

            OneNoteProxy.Instance.CommitAllModifiedPages(oneNoteApp, pageContent => pageContent.PageType == OneNoteProxy.PageType.Bible, null, null);

            return result;
        }

        public static BibleParallelTranslationConnectionResult AddParallelBible(Application oneNoteApp, string moduleShortName, string notebookDirectory, ICustomLogger logger)
        {
            if (string.IsNullOrEmpty(SettingsManager.Instance.GetValidSupplementalBibleNotebookId(oneNoteApp, true))
                || SettingsManager.Instance.SupplementalBibleModules.Count == 0)
                throw new NotConfiguredException();

            BibleParallelTranslationConnectionResult result = null;

            using (var bibleTranslationManager = new BibleParallelTranslationManager(oneNoteApp,
                SettingsManager.Instance.SupplementalBibleModules.First(), moduleShortName,
                SettingsManager.Instance.NotebookId_SupplementalBible))
            {
                if (bibleTranslationManager.BaseModuleInfo.Type == ModuleType.Strong)
                {
                    DictionaryManager.AddDictionary(oneNoteApp, moduleShortName, notebookDirectory);                    
                }

                bibleTranslationManager.Logger = logger;
                result = bibleTranslationManager.AddParallelTranslation();
            }


            SettingsManager.Instance.SupplementalBibleModules.Add(moduleShortName);
            SettingsManager.Instance.Save();

            BibleParallelTranslationManager.MergeModuleWithMainBible(moduleShortName);

            return result;
        }

        public static void CloseSupplementalBible(Application oneNoteApp)
        {
            oneNoteApp.CloseNotebook(SettingsManager.Instance.NotebookId_SupplementalBible);

            foreach (var parallelModuleName in SettingsManager.Instance.SupplementalBibleModules)
            {
                BibleParallelTranslationManager.RemoveBookAbbreviationsFromMainBible(parallelModuleName);
                var moduleInfo = ModulesManager.GetModuleInfo(parallelModuleName);
                if (moduleInfo.Type == ModuleType.Strong)
                {
                    DictionaryManager.RemoveDictionary(oneNoteApp, parallelModuleName);
                }
            }

            SettingsManager.Instance.SupplementalBibleModules.Clear();
            SettingsManager.Instance.NotebookId_SupplementalBible = null;
            SettingsManager.Instance.Save();
        }

        public enum RemoveResult
        {
            RemoveLastModule,
            RemoveSupplementalBible
        }

        public static RemoveResult RemoveLastSupplementalBibleModule(Application oneNoteApp, ICustomLogger logger)
        {
            if (SettingsManager.Instance.SupplementalBibleModules.Count <= 1)
            {
                CloseSupplementalBible(oneNoteApp);
                return RemoveResult.RemoveSupplementalBible;
            }
            else
            {
                using (var bibleTranslationManager = new BibleParallelTranslationManager(oneNoteApp,
                   SettingsManager.Instance.SupplementalBibleModules.First(), SettingsManager.Instance.SupplementalBibleModules.Last(),
                   SettingsManager.Instance.NotebookId_SupplementalBible))
                {
                    bibleTranslationManager.Logger = logger;
                    bibleTranslationManager.RemoveLastParallelTranslation();
                }

                string lastSupplementalBibleModuleName = SettingsManager.Instance.SupplementalBibleModules.Last();
                BibleParallelTranslationManager.RemoveBookAbbreviationsFromMainBible(lastSupplementalBibleModuleName);

                SettingsManager.Instance.SupplementalBibleModules.RemoveAt(SettingsManager.Instance.SupplementalBibleModules.Count - 1);
                SettingsManager.Instance.Save();

                DictionaryManager.RemoveDictionary(oneNoteApp, lastSupplementalBibleModuleName);

                return RemoveResult.RemoveLastModule;
            }
        }

        public static Dictionary<string, string> IndexStrongDictionary(Application oneNoteApp, ModuleInfo strongModuleInfo, ICustomLogger logger)
        {
            var result = new Dictionary<string, string>();
            var dictionaryModuleInfo = SettingsManager.Instance.DictionariesModules.FirstOrDefault(m => m.ModuleName == strongModuleInfo.ShortName);
            if (dictionaryModuleInfo != null)
            {
                XmlNamespaceManager xnm;
                var sectionGroupDoc = OneNoteUtils.GetHierarchyElement(oneNoteApp, dictionaryModuleInfo.SectionId, HierarchyScope.hsPages, out xnm);

                var sectionsEl = sectionGroupDoc.Root.XPathSelectElements("one:Section", xnm);
                if (sectionsEl.Count() > 0)
                {
                    foreach (var sectionEl in sectionsEl)
                    {
                        IndexStrongSection(oneNoteApp, sectionEl, result, logger, xnm);
                    }
                }
                else
                    IndexStrongSection(oneNoteApp, sectionGroupDoc.Root, result, logger, xnm); 
            }

            return result;
        }

        private static void IndexStrongSection(Application oneNoteApp, XElement sectionEl, Dictionary<string, string> result, ICustomLogger logger, XmlNamespaceManager xnm)
        {
            string sectionName = (string)sectionEl.Attribute("name");

            foreach (var pageEl in sectionEl.XPathSelectElements("one:Page", xnm))
            {
                var pageId = (string)pageEl.Attribute("ID");
                var pageDoc = OneNoteUtils.GetPageContent(oneNoteApp, pageId, out xnm);

                var tableEl = NotebookGenerator.GetPageTable(pageDoc, xnm);

                foreach (var termTextEl in tableEl.XPathSelectElements("one:Row/one:Cell[1]/one:OEChildren/one:OE/one:T", xnm))
                {
                    var termName = StringUtils.GetText(termTextEl.Value);
                    var termTextElementId = (string)termTextEl.Parent.Attribute("objectID");
                    result.Add(termName, OneNoteProxy.Instance.GenerateHref(oneNoteApp, pageId, termTextElementId));

                    if (logger != null)
                        logger.LogMessage(termName);
                }
            }
        }

        private static List<Exception> LinkdMainBibleAndSupplementalVerses(Application oneNoteApp, SimpleVersePointer baseVersePointer,
            SimpleVerse parallelVerse, BibleIteratorArgs bibleIteratorArgs, bool isStrong, Dictionary<string, string> strongTermLinksCache, 
            string alphabet, XmlNamespaceManager xnm, XNamespace nms)
        {
            var result = new List<Exception>();

            var baseBibleObjectsSearchResult = HierarchySearchManager.GetHierarchyObject(oneNoteApp,
                    SettingsManager.Instance.NotebookId_Bible, parallelVerse.ToVersePointer(SettingsManager.Instance.CurrentModule), true);

            if (baseBibleObjectsSearchResult.ResultType != HierarchySearchManager.HierarchySearchResultType.Successfully
                || baseBibleObjectsSearchResult.HierarchyStage != HierarchySearchManager.HierarchyStage.ContentPlaceholder)
                throw new ParallelVerseNotFoundException(parallelVerse, BaseVersePointerException.Severity.Error);

            var baseVerseEl = OneNoteUtils.NormalizeTextElement(
                                    HierarchySearchManager.FindVerse(bibleIteratorArgs.ChapterDocument, false, baseVersePointer.Verse, xnm));
            var baseChapterPageId = (string)bibleIteratorArgs.ChapterDocument.Root.Attribute("ID").Value;
            var baseVerseElementId = (string)baseVerseEl.Parent.Attribute("objectID").Value;

            LinkMainBibleVersesToSupplementalBibleVerse(oneNoteApp, baseChapterPageId, baseVerseElementId, parallelVerse, baseBibleObjectsSearchResult, xnm, nms);
            LinkSupplementalBibleVerseToMainBibleVerse(oneNoteApp, baseVersePointer, baseVerseEl, baseBibleObjectsSearchResult);

            if (isStrong)
            {
                baseVerseEl.Value = ProcessStrongVerse(baseVerseEl.Value, strongTermLinksCache, alphabet, result);                
            }

            return result;
        }

        private static string ProcessStrongVerse(string verseText, Dictionary<string, string> strongTermLinksCache, string alphabet, List<Exception> errors)
        {
            int cursorPosition = StringUtils.GetNextIndexOfDigit(verseText, null);
            int temp, htmlBreakIndex = -1;
            string strongNumber;

            while (cursorPosition > -1)
            {
                strongNumber = StringUtils.GetNextString(verseText, cursorPosition - 1, new SearchMissInfo(null, SearchMissInfo.MissMode.CancelOnMissFound), alphabet,
                                                                    out temp, out htmlBreakIndex, StringSearchIgnorance.None, StringSearchMode.SearchNumber);
                if (!string.IsNullOrEmpty(strongNumber))
                {
                    string prefix = StringUtils.GetPrevString(verseText, cursorPosition, new SearchMissInfo(null, SearchMissInfo.MissMode.CancelOnMissFound), alphabet,
                                                                    out temp, out temp, StringSearchIgnorance.None, StringSearchMode.SearchFirstChar);
                    if (!string.IsNullOrEmpty(prefix) && prefix.Length == 1 && StringUtils.IsCharAlphabetical(prefix[0], alphabet))
                    {
                        string strongTerm = prefix + strongNumber;
                        string link = string.Format("<a href=\"{0}\">{1}</a>", strongTermLinksCache[strongTerm], strongTerm); //добавить <sup>
                        if (strongTermLinksCache.ContainsKey(strongTerm))
                        {
                            verseText = string.Concat(verseText.Substring(0, cursorPosition), link, verseText.Substring(htmlBreakIndex));
                            htmlBreakIndex += link.Length;
                        }
                        else
                            errors.Add(new Exception(string.Format("There is no strongTermName '{0}' in strongTermLinksCache", strongTerm)));
                    }
                }

                cursorPosition = StringUtils.GetNextIndexOfDigit(verseText, htmlBreakIndex);
            }

            return verseText;
        }

        private static void LinkSupplementalBibleVerseToMainBibleVerse(Application oneNoteApp, SimpleVersePointer baseVersePointer, XElement baseVerseEl, HierarchySearchManager.HierarchySearchResult baseBibleObjectsSearchResult)
        {
            int textBreakIndex, htmlBreakIndex;
            var baseVerseNumber = StringUtils.GetNextString(baseVerseEl.Value, -1, new SearchMissInfo(0, SearchMissInfo.MissMode.CancelOnMissFound), out textBreakIndex, out htmlBreakIndex);

            if (baseVerseNumber != baseVersePointer.Verse.ToString())
                throw new InvalidOperationException(
                    string.Format("baseVerseNumber != baseVersePointer (baseVerseNumber = '{0}', baseVersePointer = '{1}')", baseVerseNumber, baseVersePointer));

            string linkToParallelVerse = OneNoteUtils.GenerateHref(oneNoteApp, baseVerseNumber,
                baseBibleObjectsSearchResult.HierarchyObjectInfo.PageId, baseBibleObjectsSearchResult.HierarchyObjectInfo.ContentObjectId);

            baseVerseEl.Value = string.Format("{0} {1}", linkToParallelVerse, baseVerseEl.Value.Substring(htmlBreakIndex + 1));
        }

        private static void LinkMainBibleVersesToSupplementalBibleVerse(Application oneNoteApp, string baseChapterPageId, string baseVerseElementId, 
            SimpleVerse parallelVerse, HierarchySearchManager.HierarchySearchResult baseBibleObjectsSearchResult, XmlNamespaceManager xnm, XNamespace nms)
        {           
            if (parallelVerse.PartIndex.GetValueOrDefault(0) == 0 && !parallelVerse.IsEmpty && !string.IsNullOrEmpty(parallelVerse.VerseContent))  // если PartIndex > 0, значит этот стих мы уже привязали
            {
                var parallelChapterPageDoc = PrepareMainBibleTable(oneNoteApp, baseBibleObjectsSearchResult.HierarchyObjectInfo.PageId);

                string linkToBaseVerse = OneNoteUtils.GenerateHref(oneNoteApp, SettingsManager.Instance.SupplementalBibleLinkName, baseChapterPageId, baseVerseElementId);

                foreach (var parallelVerseElementId in baseBibleObjectsSearchResult.HierarchyObjectInfo.GetAllObjectsIds())
                {                    
                    var cell = parallelChapterPageDoc.Content.Root
                                    .XPathSelectElement(string.Format("//one:OE[@objectID='{0}']", parallelVerseElementId), xnm).Parent.Parent;
                    var row = cell.Parent;
                    if (row.Elements().Count() == 3)
                        row.Elements().Last().XPathSelectElement("one:OEChildren/one:OE/one:T", xnm).Value = linkToBaseVerse;
                    else
                        row.Add(NotebookGenerator.GetCell(linkToBaseVerse, string.Empty, nms));
                }
            }
        }

        private static OneNoteProxy.PageContent PrepareMainBibleTable(Application oneNoteApp, string mainBibleChapterPageId)
        {
            var parallelChapterPageDoc = OneNoteProxy.Instance.GetPageContent(oneNoteApp, mainBibleChapterPageId, OneNoteProxy.PageType.Bible);
            var parallelBibleTableElement = NotebookGenerator.GetPageTable(parallelChapterPageDoc.Content, parallelChapterPageDoc.Xnm);

            var columnsCount = parallelBibleTableElement.XPathSelectElements("one:Columns/one:Column", parallelChapterPageDoc.Xnm).Count();
            if (columnsCount == 2)
                NotebookGenerator.AddColumnToTable(parallelBibleTableElement, NotebookGenerator.MinimalCellWidth, parallelChapterPageDoc.Xnm);
            parallelChapterPageDoc.WasModified = true;

            return parallelChapterPageDoc;
        }               
     

        private static void GenerateChapterPage(Application oneNoteApp, BibleChapterContent chapter, string bookSectionId,
           ModuleInfo moduleInfo, BibleBookInfo bibleBookInfo, ModuleBibleInfo bibleInfo)
        {
            string chapterPageName = string.Format(moduleInfo.BibleStructure.ChapterSectionNameTemplate, chapter.Index, bibleBookInfo.Name);

            XmlNamespaceManager xnm;
            var currentChapterDoc = NotebookGenerator.AddPage(oneNoteApp, bookSectionId, chapterPageName, 1, bibleInfo.Content.Locale, out xnm);

            var currentTableElement = NotebookGenerator.AddTableToPage(currentChapterDoc, false, xnm, new CellInfo(SettingsManager.Instance.PageWidth_Bible));

            NotebookGenerator.AddParallelBibleTitle(currentTableElement, moduleInfo.Name, 0, bibleInfo.Content.Locale, xnm);

            foreach (var verse in chapter.Verses)
            {
                NotebookGenerator.AddVerseRowToTable(currentTableElement, string.Format("{0} {1}", verse.Index, verse.Value), 0, bibleInfo.Content.Locale);
            }

            UpdateChapterPage(oneNoteApp, currentChapterDoc, chapter.Index, bibleBookInfo);
        }       

        private static string GetCurrentSectionGroupId(Application oneNoteApp, string currentSectionGroupId, Common.ModuleInfo moduleInfo, int i)
        {
            if (string.IsNullOrEmpty(currentSectionGroupId))
                currentSectionGroupId
                    = NotebookGenerator.AddRootSectionGroupToNotebook(oneNoteApp,
                        SettingsManager.Instance.NotebookId_SupplementalBible, moduleInfo.BibleStructure.OldTestamentName).Attribute("ID").Value;
            else if (i == moduleInfo.BibleStructure.OldTestamentBooksCount)
                currentSectionGroupId
                    = NotebookGenerator.AddRootSectionGroupToNotebook(oneNoteApp,
                        SettingsManager.Instance.NotebookId_SupplementalBible, moduleInfo.BibleStructure.NewTestamentName).Attribute("ID").Value;

            return currentSectionGroupId;
        }

        private static void UpdateChapterPage(Application oneNoteApp, XDocument chapterPageDoc, int chapterIndex, BibleBookInfo bibleBookInfo)
        {
            oneNoteApp.UpdatePageContent(chapterPageDoc.ToString(), DateTime.MinValue, Constants.CurrentOneNoteSchema);            
        }       
    }
}
