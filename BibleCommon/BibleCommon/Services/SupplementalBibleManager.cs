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
using BibleCommon.Scheme;
using BibleCommon.Handlers;

namespace BibleCommon.Services
{
    public static class SupplementalBibleManager
    {
        public static void CreateSupplementalBible(Application oneNoteApp, ModuleInfo module, string notebookDirectory, ICustomLogger logger)
        {
            if (string.IsNullOrEmpty(SettingsManager.Instance.GetValidSupplementalBibleNotebookId(oneNoteApp, true)))
            {
                SettingsManager.Instance.NotebookId_SupplementalBible
                    = NotebookGenerator.CreateNotebook(oneNoteApp, Resources.Constants.SupplementalBibleName, notebookDirectory);                
            }            

            string currentSectionGroupId = null;
            string currentStrongPrefix = null;
            var moduleInfo = ModulesManager.GetModuleInfo(module.ShortName);
            var bibleInfo = ModulesManager.GetModuleBibleInfo(module.ShortName);            

            string oldTestamentName = null;
            int? oldTestamentSectionsCount = null;
            string oldTestamentStrongPrefix = null;
            string newTestamentName = null;
            int? newTestamentSectionsCount = null;
            string newTestamentStrongPrefix = null;

            GetTestamentInfo(moduleInfo, ContainerType.OldTestament, out oldTestamentName, out oldTestamentSectionsCount, out oldTestamentStrongPrefix);
            GetTestamentInfo(moduleInfo, ContainerType.NewTestament, out newTestamentName, out newTestamentSectionsCount, out newTestamentStrongPrefix);

            SettingsManager.Instance.SupplementalBibleModules.Clear();
            SettingsManager.Instance.SupplementalBibleModules.Add(new StoredModuleInfo(module.ShortName, module.Version));
            SettingsManager.Instance.Save();
            
            for (int i = 0; i < moduleInfo.BibleStructure.BibleBooks.Count; i++)
            {
                var bibleBookInfo = moduleInfo.BibleStructure.BibleBooks[i];

                bibleBookInfo.SectionName = NotebookGenerator.GetBibleBookSectionName(bibleBookInfo.Name, i, (oldTestamentSectionsCount ?? newTestamentSectionsCount).Value);

                currentSectionGroupId = GetCurrentSectionGroupId(oneNoteApp, currentSectionGroupId, 
                    oldTestamentName, oldTestamentSectionsCount, newTestamentName, newTestamentSectionsCount, i);

                if (moduleInfo.Type == Common.ModuleType.Strong)
                {
                    currentStrongPrefix = GetStrongPrefix(i + 1, (oldTestamentSectionsCount ?? newTestamentSectionsCount).Value, oldTestamentStrongPrefix, newTestamentStrongPrefix);
                }

                var bookSectionId = NotebookGenerator.AddSection(oneNoteApp, currentSectionGroupId, bibleBookInfo.SectionName);

                var bibleBook = bibleInfo.Books.FirstOrDefault(book => book.Index == bibleBookInfo.Index);
                if (bibleBook == null)
                    throw new Exception("Manifest.xml has Bible books that do not exist in bible.xml");

                foreach (var chapter in bibleBook.Chapters)
                {
                    if (logger != null)
                        logger.LogMessage("{0} '{1} {2}'", BibleCommon.Resources.Constants.ProcessChapter, bibleBookInfo.Name, chapter.Index);

                    GenerateChapterPage(oneNoteApp, chapter, bookSectionId, moduleInfo, bibleBookInfo, bibleInfo, currentStrongPrefix);
                }
            }

            oneNoteApp.SyncHierarchy(SettingsManager.Instance.NotebookId_SupplementalBible);            
        }

        private static string GetStrongPrefix(int bookIndex, int oldTestamentBooksCount, string oldTestamentStrongPrefix, string newTestamentStrongPrefix)
        {
            return bookIndex > oldTestamentBooksCount ? newTestamentStrongPrefix : (oldTestamentStrongPrefix ?? newTestamentStrongPrefix);
        }

        private static void GetTestamentInfo(ModuleInfo moduleInfo, ContainerType type, out string testamentName, out int? testamentSectionsCount, out string strongPrefix)
        {
            testamentName = null;
            testamentSectionsCount = null;
            strongPrefix = null;

            var testamentSectionGroup = moduleInfo.GetNotebook(ContainerType.Bible).SectionGroups.FirstOrDefault(s => s.Type == type);
            if (testamentSectionGroup != null)
            {
                testamentName = testamentSectionGroup.Name;
                testamentSectionsCount = testamentSectionGroup.SectionsCount;
                strongPrefix = testamentSectionGroup.StrongPrefix;
            }
        }

        private static void UnlockSupplementalBible(Application oneNoteApp)
        {
            try
            {
                OneNoteLocker.UnlockBible(oneNoteApp);
                //OneNoteLocker.UnlockSupplementalBible(oneNoteApp);  // пока вроде как это не надо, так как данный метод вызывается только при создании спр Библии
            }
            catch (NotSupportedException)
            {
                //todo: log it
            }
        }

        public static BibleParallelTranslationConnectionResult LinkSupplementalBibleWithPrimaryBible(Application oneNoteApp, int supplementalModuleIndex,
            Dictionary<string, string> strongTermLinksCache, ICustomLogger logger)
        {
            if (supplementalModuleIndex != 0)
                throw new NotSupportedException("supplementalModuleIndex != 0");            

            if (string.IsNullOrEmpty(SettingsManager.Instance.GetValidSupplementalBibleNotebookId(oneNoteApp, true)) 
                || SettingsManager.Instance.SupplementalBibleModules.Count == 0)
                throw new NotConfiguredException("Supplemental Bible does not exists.");

            XmlNamespaceManager xnm = OneNoteUtils.GetOneNoteXNM();
            var nms = XNamespace.Get(Constants.OneNoteXmlNs);

            string supplementalModuleShortName = SettingsManager.Instance.SupplementalBibleModules[supplementalModuleIndex].ModuleName;            

            BibleParallelTranslationManager.MergeModuleWithMainBible(supplementalModuleShortName);

            UnlockSupplementalBible(oneNoteApp);

            var isOneNote2010 = OneNoteUtils.IsOneNote2010Cached(oneNoteApp);

            BibleParallelTranslationConnectionResult result;
            using (var bibleTranslationManager = new BibleParallelTranslationManager(oneNoteApp,
                            supplementalModuleShortName, SettingsManager.Instance.ModuleShortName,
                            SettingsManager.Instance.NotebookId_SupplementalBible))
            {
                if (bibleTranslationManager.BaseModuleInfo.Type == Common.ModuleType.Strong)
                    if (strongTermLinksCache == null)
                        throw new ArgumentNullException("strongTermLinksCache");                

                bibleTranslationManager.Logger = logger;                               

                var linkResult = new List<Exception>();

                result = bibleTranslationManager.IterateBaseBible(
                    (chapterPageDoc, chapterPointer) =>
                    {
                        OneNoteProxy.Instance.CommitAllModifiedPages(oneNoteApp, pageContent => pageContent.PageType == OneNoteProxy.PageType.Bible, null, null);

                        int styleIndex = QuickStyleManager.AddQuickStyleDef(chapterPageDoc, QuickStyleManager.StyleForStrongName, QuickStyleManager.PredefinedStyles.GrayHyperlink, xnm);

                        return new BibleIteratorArgs() { ChapterDocument = chapterPageDoc, StrongStyleIndex = styleIndex };
                    }, true, true,
                    (baseVersePointer, parallelVerse, bibleIteratorArgs) =>
                    {
                        if (!parallelVerse.IsEmpty)
                        {
                            linkResult.AddRange(
                                LinkdPrimaryBibleAndSupplementalVerses(oneNoteApp, baseVersePointer, parallelVerse, bibleIteratorArgs,
                                            bibleTranslationManager.BaseModuleInfo.Type == Common.ModuleType.Strong, strongTermLinksCache,
                                            bibleTranslationManager.BaseModuleInfo.ShortName,
                                            bibleTranslationManager.BaseModuleInfo.BibleStructure.Alphabet, isOneNote2010, xnm, nms));
                        }
                    });

                result.Errors.AddRange(linkResult);
            }

            OneNoteProxy.Instance.CommitAllModifiedPages(oneNoteApp, pageContent => pageContent.PageType == OneNoteProxy.PageType.Bible, null, null);

            return result;
        }

        public static BibleParallelTranslationConnectionResult AddParallelBible(Application oneNoteApp, ModuleInfo module, 
                Dictionary<string, string> strongTermLinksCache, ICustomLogger logger)
        {
            if (string.IsNullOrEmpty(SettingsManager.Instance.GetValidSupplementalBibleNotebookId(oneNoteApp, true))
                || SettingsManager.Instance.SupplementalBibleModules.Count == 0)
                throw new NotConfiguredException();

            SettingsManager.Instance.SupplementalBibleModules.Add(new StoredModuleInfo(module.ShortName, module.Version));
            SettingsManager.Instance.Save();
            
            BibleParallelTranslationManager.MergeModuleWithMainBible(module.ShortName);

            string oldTestamentName = null;
            int? oldTestamentSectionsCount = null;
            string oldTestamentStrongPrefix = null;
            string newTestamentName = null;
            int? newTestamentSectionsCount = null;
            string newTestamentStrongPrefix = null;            

            BibleParallelTranslationConnectionResult result = null;
            XmlNamespaceManager xnm = OneNoteUtils.GetOneNoteXNM();
            var linkResult = new List<Exception>();
            var isOneNote2010 = OneNoteUtils.IsOneNote2010Cached(oneNoteApp);

            using (var bibleTranslationManager = new BibleParallelTranslationManager(oneNoteApp,
                SettingsManager.Instance.SupplementalBibleModules.First().ModuleName, module.ShortName,
                SettingsManager.Instance.NotebookId_SupplementalBible))
            {
                UnlockSupplementalBible(oneNoteApp);

                GetTestamentInfo(bibleTranslationManager.ParallelModuleInfo, ContainerType.OldTestament, out oldTestamentName, out oldTestamentSectionsCount, out oldTestamentStrongPrefix);
                GetTestamentInfo(bibleTranslationManager.ParallelModuleInfo, ContainerType.NewTestament, out newTestamentName, out newTestamentSectionsCount, out newTestamentStrongPrefix);                

                bibleTranslationManager.Logger = logger;
                result = bibleTranslationManager.IterateBaseBible(
                    (chapterPageDoc, chapterPointer) =>
                    {
                        UpdateSupplementalModulesMetadata(oneNoteApp, chapterPageDoc, chapterPointer, module, xnm);

                        var tableEl = NotebookGenerator.GetPageTable(chapterPageDoc, xnm);
                        var bibleIndex = NotebookGenerator.AddColumnToTable(tableEl, SettingsManager.Instance.PageWidth_Bible, xnm);
                        NotebookGenerator.AddParallelBibleTitle(chapterPageDoc, tableEl, 
                            bibleTranslationManager.ParallelModuleInfo.DisplayName, bibleIndex, bibleTranslationManager.ParallelModuleInfo.Locale, xnm);

                        int styleIndex = QuickStyleManager.AddQuickStyleDef(chapterPageDoc, QuickStyleManager.StyleForStrongName, QuickStyleManager.PredefinedStyles.GrayHyperlink, xnm);

                        var strongPrefix = bibleTranslationManager.ParallelModuleInfo.Type == Common.ModuleType.Strong 
                            ? GetStrongPrefix(chapterPointer.BookIndex, (oldTestamentSectionsCount ?? newTestamentSectionsCount).Value, oldTestamentStrongPrefix, newTestamentStrongPrefix)
                            : null;

                        return new BibleIteratorArgs() { BibleIndex = bibleIndex, TableElement = tableEl, StrongStyleIndex = styleIndex, StrongPrefix = strongPrefix };
                    },               
                    true, true,
                    (baseVersePointer, parallelVerse, bibleIteratorArgs) =>
                    {
                        if (!parallelVerse.IsEmpty)
                        {
                            if (bibleTranslationManager.ParallelModuleInfo.Type == Common.ModuleType.Strong)
                            {
                                parallelVerse.VerseContent = ProcessStrongVerse(parallelVerse.VerseContent, strongTermLinksCache,
                                    bibleTranslationManager.ParallelModuleShortName,
                                    bibleTranslationManager.ParallelModuleInfo.BibleStructure.Alphabet, isOneNote2010, ref linkResult);
                            }

                            var cell = NotebookGenerator.AddParallelVerseRowToBibleTable(bibleIteratorArgs.TableElement, parallelVerse,
                                bibleIteratorArgs.BibleIndex, baseVersePointer, bibleTranslationManager.ParallelModuleInfo.Locale, xnm);

                            if (bibleTranslationManager.ParallelModuleInfo.Type == Common.ModuleType.Strong)
                            {
                                QuickStyleManager.SetQuickStyleDefForCell(cell, bibleIteratorArgs.StrongStyleIndex, xnm);
                            }
                        }
                    });
            }

            result.Errors.AddRange(linkResult);            

            return result;
        }

        private static void UpdateSupplementalModulesMetadata(Application oneNoteApp, XDocument chapterPageDoc, SimpleVersePointer chapterPointer, ModuleInfo module,
            XmlNamespaceManager xnm)
        {
            var supplementalModulesMetadata = OneNoteUtils.GetPageMetaData(oneNoteApp, chapterPageDoc.Root, Consts.Constants.EmbeddedSupplementalModulesKey, xnm);
            if (string.IsNullOrEmpty(supplementalModulesMetadata))
                throw new InvalidOperationException(string.Format("Chapter page metadata was not found: {0}", chapterPointer));

            var supplementalModulesInfo = EmbeddedModuleInfo.Deserialize(supplementalModulesMetadata);
            supplementalModulesInfo.Add(new EmbeddedModuleInfo(module.ShortName, module.Version, supplementalModulesInfo.Count));

            OneNoteUtils.UpdatePageMetaData(oneNoteApp, chapterPageDoc.Root, Consts.Constants.EmbeddedSupplementalModulesKey, 
                EmbeddedModuleInfo.Serialize(supplementalModulesInfo), xnm);
        }

        public static void CloseSupplementalBible(Application oneNoteApp)
        {            
            OneNoteUtils.CloseNotebookSafe(oneNoteApp, SettingsManager.Instance.NotebookId_SupplementalBible);

            foreach (var parallelModuleName in SettingsManager.Instance.SupplementalBibleModules)
            {
                BibleParallelTranslationManager.RemoveBookAbbreviationsFromMainBible(parallelModuleName.ModuleName);
                var moduleInfo = ModulesManager.GetModuleInfo(parallelModuleName.ModuleName);
                if (moduleInfo.Type == Common.ModuleType.Strong)
                {
                    DictionaryManager.RemoveDictionary(oneNoteApp, parallelModuleName.ModuleName);
                }
            }

            SettingsManager.Instance.SupplementalBibleModules.Clear();
            SettingsManager.Instance.NotebookId_SupplementalBible = null;
            SettingsManager.Instance.Save();
        }

        public enum RemoveResult
        {
            RemoveModule,
            RemoveSupplementalBible
        }

        public static RemoveResult RemoveSupplementalBibleModule(Application oneNoteApp, string moduleShortName, ICustomLogger logger)
        {
            if (SettingsManager.Instance.SupplementalBibleModules.Count <= 1)
            {
                CloseSupplementalBible(oneNoteApp);
                return RemoveResult.RemoveSupplementalBible;
            }
            else
            {
                var storedModuleInfo = SettingsManager.Instance.SupplementalBibleModules.FirstOrDefault(m => m.ModuleName == moduleShortName);

                if (storedModuleInfo == null)
                    throw new ArgumentException(string.Format("Module '{0}' can not be found in Supplemental Bible", moduleShortName));

                var moduleInfo = ModulesManager.GetModuleInfo(moduleShortName);
                if (moduleInfo.Type == Common.ModuleType.Strong)
                    DictionaryManager.RemoveDictionary(oneNoteApp, moduleShortName);

                BibleParallelTranslationManager.RemoveBookAbbreviationsFromMainBible(moduleShortName);

                SettingsManager.Instance.SupplementalBibleModules.Remove(storedModuleInfo);
                SettingsManager.Instance.Save();

                using (var bibleTranslationManager = new BibleParallelTranslationManager(oneNoteApp,
                   SettingsManager.Instance.SupplementalBibleModules.First().ModuleName, moduleShortName,
                   SettingsManager.Instance.NotebookId_SupplementalBible))
                {
                    bibleTranslationManager.Logger = logger;
                    bibleTranslationManager.RemoveParallelTranslation(moduleShortName);
                }                

                return RemoveResult.RemoveModule;
            }
        }       

        // перед обновлением страницы Библии со стронгом нужно обязательно вызывать этот метод, иначе все ссылки станут синими
        public static void UpdatePageXmlForStrongBible(XDocument pageDoc, bool isOneNote2010)
        {
            if (!isOneNote2010)
                return;

            XmlNamespaceManager xnm = OneNoteUtils.GetOneNoteXNM();
            var styleEl = pageDoc.Root.XPathSelectElement(string.Format("one:QuickStyleDef[@name='{0}']", QuickStyleManager.StyleForStrongName), xnm);
            if (styleEl != null)  // значит видимо есть Библия Стронга на текущей странице
            {
                string searchTemplate = "</a>";
                XNamespace nms = XNamespace.Get(Constants.OneNoteXmlNs);                

                foreach (var textEl in pageDoc.Root.XPathSelectElements(string.Format("//one:OE[@quickStyleIndex='{0}']/one:T", (string)styleEl.Attribute("index")), xnm))
                {
                    OneNoteUtils.NormalizeTextElement(textEl);
                    int firstLinkEndIndex = textEl.Value.IndexOf(searchTemplate);

                    if (firstLinkEndIndex != -1)
                    {
                        firstLinkEndIndex = firstLinkEndIndex + searchTemplate.Length;
                        var nextTagStartIndex = textEl.Value.IndexOf("<", firstLinkEndIndex);
                        var nextTagEndIndex = textEl.Value.IndexOf("</", firstLinkEndIndex);

                        if (nextTagEndIndex == nextTagStartIndex)
                        {
                            firstLinkEndIndex = textEl.Value.IndexOf(">", nextTagEndIndex + 1) + 1;                            
                        }

                        string firstLink = textEl.Value.Substring(0, firstLinkEndIndex);

                        textEl.AddBeforeSelf(new XElement(nms + "T",
                                               new XCData(firstLink))
                                             );

                        textEl.Value = string.Format(" {0}", textEl.Value.Substring(firstLinkEndIndex));
                    }
                }
            }
        }      

        private static List<Exception> LinkdPrimaryBibleAndSupplementalVerses(Application oneNoteApp, SimpleVersePointer baseVersePointer,
            SimpleVerse parallelVerse, BibleIteratorArgs bibleIteratorArgs, bool isStrong, Dictionary<string, string> strongTermLinksCache, 
            string strongModuleShortName, string alphabet, bool isOneNote2010, XmlNamespaceManager xnm, XNamespace nms)
        {
            var result = new List<Exception>();            

            var primaryBibleObjectsSearchResult = HierarchySearchManager.GetHierarchyObject(oneNoteApp,
                    SettingsManager.Instance.NotebookId_Bible, parallelVerse.ToVersePointer(SettingsManager.Instance.CurrentModuleCached), HierarchySearchManager.FindVerseLevel.AllVerses);

            if (primaryBibleObjectsSearchResult.ResultType != HierarchySearchManager.HierarchySearchResultType.Successfully
                || primaryBibleObjectsSearchResult.HierarchyStage != HierarchySearchManager.HierarchyStage.ContentPlaceholder)
                throw new VerseNotFoundException(parallelVerse, SettingsManager.Instance.ModuleShortName, BaseVersePointerException.Severity.Error);

            VerseNumber? baseVerseNumber;
            string verseTextWithoutNumber;
            var baseVerseEl = OneNoteUtils.NormalizeTextElement(
                                    HierarchySearchManager.FindVerse(bibleIteratorArgs.ChapterDocument, false, baseVersePointer.Verse, xnm,
                                    out baseVerseNumber, out verseTextWithoutNumber));            
                
            var baseChapterPageId = (string)bibleIteratorArgs.ChapterDocument.Root.Attribute("ID");
            var baseVerseElementId = (string)baseVerseEl.Parent.Attribute("objectID");            

            LinkMainBibleVersesToSupplementalBibleVerse(oneNoteApp, baseChapterPageId, baseVerseElementId, parallelVerse, primaryBibleObjectsSearchResult, xnm, nms);
            LinkSupplementalBibleVerseToMainBibleVerseAndToStrongDictionary(oneNoteApp, baseVersePointer, baseVerseEl, baseVerseNumber, verseTextWithoutNumber, primaryBibleObjectsSearchResult, 
                isStrong, bibleIteratorArgs.StrongStyleIndex, strongTermLinksCache, strongModuleShortName, alphabet, isOneNote2010, ref result, nms);

            return result;
        }

        private static string ProcessStrongVerse(string verseText, Dictionary<string, string> strongTermLinksCache, 
            string strongModuleShortName, string alphabet, bool isOneNote2010, ref List<Exception> errors)
        {
            int cursorPosition = StringUtils.GetNextIndexOfDigit(verseText, null);
            int temp, htmlBreakIndex = -1;
            string strongNumber;

            if (isOneNote2010)
            {
                var verseParts = StringUtils.GetText(verseText).Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                for (int i = 0; i < verseParts.Length; i++)
                {
                    verseParts[i] = string.Format("<span style='color:#000000'>{0}</span>", verseParts[i]);
                }
                verseText = string.Join(" ", verseParts);
            }
            
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
                        string strongTerm = string.Format("{0}{1:0000}", prefix, int.Parse(strongNumber));
                        if (strongTermLinksCache.ContainsKey(strongTerm))
                        {
                            var termLink = new DictionaryTermLink(strongTermLinksCache[strongTerm]).Href;
                            string link = string.Format("<a href=\"{0}\"><span style='vertical-align:super;'>{1}</span></a>",
                                SettingsManager.Instance.UseMiddleStrongLinks || string.IsNullOrEmpty(termLink)
                                    ? NavigateToStrongHandler.GetCommandUrlStatic(strongTerm, strongModuleShortName)
                                    : termLink,
                                strongTerm);

                            verseText = string.Concat(verseText.Substring(0, cursorPosition - 1), link, verseText.Substring(htmlBreakIndex));

                            htmlBreakIndex += link.Length - strongNumber.Length - 1;
                        }
                        else
                            errors.Add(new Exception(string.Format("There is no strongTermName '{0}' in strongTermLinksCache", strongTerm)));
                    }
                }

                cursorPosition = StringUtils.GetNextIndexOfDigit(verseText, htmlBreakIndex);
            }

            return verseText;
        }

        private static void LinkSupplementalBibleVerseToMainBibleVerseAndToStrongDictionary(Application oneNoteApp, 
            SimpleVersePointer baseVersePointer, XElement baseVerseEl, VerseNumber? baseVerseNumber, string verseTextWithoutNumber,
            HierarchySearchManager.HierarchySearchResult primaryBibleObjectsSearchResult,
            bool isStrong, int strongStyleIndex, Dictionary<string, string> strongTermLinksCache, string strongModuleShortName, string alphabet, bool isOneNote2010,
            ref List<Exception> result, XNamespace nms)
        {
            if (baseVersePointer.VerseNumber != baseVerseNumber)            
                result.Add(
                    new InvalidOperationException(
                        string.Format("baseVerseNumber != baseVersePointer (baseVerseNumber = '{0}', baseVersePointer = '{1}')", baseVerseNumber, baseVersePointer)));

            string linkToParallelVerse = OneNoteUtils.GetOrGenerateHref(oneNoteApp, baseVerseNumber.ToString(),
                primaryBibleObjectsSearchResult.HierarchyObjectInfo.VerseInfo.ObjectHref,
                primaryBibleObjectsSearchResult.HierarchyObjectInfo.PageId, primaryBibleObjectsSearchResult.HierarchyObjectInfo.VerseContentObjectId);

            string versePart = verseTextWithoutNumber;

            if (isStrong)
            {
                if (isOneNote2010)
                    baseVerseEl.Parent.SetAttributeValue("quickStyleIndex", strongStyleIndex);
                versePart = ProcessStrongVerse(versePart, strongTermLinksCache, strongModuleShortName, alphabet, isOneNote2010, ref result);
            }

            baseVerseEl.Value = string.Format("{0}<span> </span>{1}", linkToParallelVerse, versePart);
        }

        private static void LinkMainBibleVersesToSupplementalBibleVerse(Application oneNoteApp, string baseChapterPageId, string baseVerseElementId, 
            SimpleVerse parallelVerse, HierarchySearchManager.HierarchySearchResult baseBibleObjectsSearchResult, XmlNamespaceManager xnm, XNamespace nms)
        {           
            if (parallelVerse.PartIndex.GetValueOrDefault(0) == 0 && !parallelVerse.IsEmpty && !string.IsNullOrEmpty(parallelVerse.VerseContent))  // если PartIndex > 0, значит этот стих мы уже привязали
            {
                var parallelChapterPageDoc = PrepareMainBibleTable(oneNoteApp, baseBibleObjectsSearchResult.HierarchyObjectInfo.PageId);

                string linkToBaseVerse = string.Format("<font size='2pt'>{0}</font>",
                                            OneNoteUtils.GenerateHref(oneNoteApp, SettingsManager.Instance.SupplementalBibleLinkName, baseChapterPageId, baseVerseElementId));

                foreach (var parallelVerseElementId in baseBibleObjectsSearchResult.HierarchyObjectInfo.GetAllObjectsIds())
                {                    
                    var bibleCell = parallelChapterPageDoc.Content.Root
                                    .XPathSelectElement(string.Format("//one:OE[@objectID='{0}']", parallelVerseElementId.ObjectId), xnm).Parent.Parent;
                    var row = bibleCell.Parent;
                    XElement sCell = null;
                    if (row.Elements().Count() == 3)
                    {
                        sCell = row.Elements().Last();
                        sCell.XPathSelectElement("one:OEChildren/one:OE/one:T", xnm).Value = linkToBaseVerse;
                    }
                    else
                    {
                        sCell = NotebookGenerator.GetCell(linkToBaseVerse, string.Empty, nms);
                        row.Add(sCell);
                    }

                    sCell.XPathSelectElement("one:OEChildren/one:OE", xnm).SetAttributeValue("alignment", "center");
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


        private static void GenerateChapterPage(Application oneNoteApp, CHAPTER chapter, string bookSectionId,
           ModuleInfo moduleInfo, BibleBookInfo bibleBookInfo, XMLBIBLE bibleInfo, string strongPrefix)
        {
            string chapterPageName = string.Format(!string.IsNullOrEmpty(bibleBookInfo.ChapterPageNameTemplate) 
                                                        ? bibleBookInfo.ChapterPageNameTemplate
                                                        : moduleInfo.BibleStructure.ChapterPageNameTemplate, 
                                                   chapter.Index, bibleBookInfo.Name);

            XmlNamespaceManager xnm;
            var currentChapterDoc = NotebookGenerator.AddPage(oneNoteApp, bookSectionId, chapterPageName, 1, moduleInfo.Locale, out xnm);

            OneNoteUtils.UpdatePageMetaData(oneNoteApp, currentChapterDoc.Root, Consts.Constants.EmbeddedSupplementalModulesKey,
                EmbeddedModuleInfo.Serialize(new List<EmbeddedModuleInfo>() { new EmbeddedModuleInfo(moduleInfo.ShortName, moduleInfo.Version, 0) }), xnm);

            var currentTableElement = NotebookGenerator.AddTableToPage(currentChapterDoc, false, xnm, new CellInfo(SettingsManager.Instance.PageWidth_Bible));

            NotebookGenerator.AddParallelBibleTitle(currentChapterDoc, currentTableElement, moduleInfo.DisplayName, 0, moduleInfo.Locale, xnm);

            foreach (var verse in chapter.Verses)
            {                
                NotebookGenerator.AddVerseRowToTable(currentTableElement, BIBLEBOOK.GetFullVerseString(verse.Index, verse.TopIndex, verse.GetValue(true, strongPrefix)), 0, moduleInfo.Locale);
            }

            UpdateChapterPage(oneNoteApp, currentChapterDoc, chapter.Index, bibleBookInfo);
        }

        private static string GetCurrentSectionGroupId(Application oneNoteApp, string currentSectionGroupId, 
            string oldTestamentName, int? oldTestamentSectionsCount, string newTestamentName, int? newTestamentSectionsCount, int i)
        {
            if (string.IsNullOrEmpty(currentSectionGroupId))
            {
                currentSectionGroupId
                    = (string)NotebookGenerator.AddRootSectionGroupToNotebook(oneNoteApp,
                        SettingsManager.Instance.NotebookId_SupplementalBible, oldTestamentName ?? newTestamentName).Attribute("ID");                
            }
            else if (i == (oldTestamentSectionsCount ?? newTestamentSectionsCount))  // если только один завет в модуле, то до сюда и не должен дойти
            {
                currentSectionGroupId
                    = (string)NotebookGenerator.AddRootSectionGroupToNotebook(oneNoteApp,
                        SettingsManager.Instance.NotebookId_SupplementalBible, newTestamentName).Attribute("ID");             
            }            

            return currentSectionGroupId;
        }
        private static void UpdateChapterPage(Application oneNoteApp, XDocument chapterPageDoc, int chapterIndex, BibleBookInfo bibleBookInfo)
        {
            oneNoteApp.UpdatePageContent(chapterPageDoc.ToString(), DateTime.MinValue, Constants.CurrentOneNoteSchema);            
        }       
    }
}
