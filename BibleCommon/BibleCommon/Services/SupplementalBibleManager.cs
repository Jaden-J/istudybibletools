﻿using System;
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
        public static void CreateSupplementalBible(ref Application oneNoteApp, ModuleInfo module, string notebookDirectory, ICustomLogger logger)
        {
            if (string.IsNullOrEmpty(SettingsManager.Instance.GetValidSupplementalBibleNotebookId(ref oneNoteApp, true)))
            {
                SettingsManager.Instance.NotebookId_SupplementalBible
                    = NotebookGenerator.CreateNotebook(ref oneNoteApp, Resources.Constants.SupplementalBibleName, notebookDirectory, Resources.Constants.SupplementalBibleName);
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

                currentSectionGroupId = GetCurrentSectionGroupId(ref oneNoteApp, currentSectionGroupId, 
                    oldTestamentName, oldTestamentSectionsCount, newTestamentName, newTestamentSectionsCount, i);

                if (moduleInfo.Type == Common.ModuleType.Strong)
                {
                    currentStrongPrefix = GetStrongPrefix(i + 1, (oldTestamentSectionsCount ?? newTestamentSectionsCount).Value, oldTestamentStrongPrefix, newTestamentStrongPrefix);
                }

                var bookSectionId = NotebookGenerator.AddSection(ref oneNoteApp, currentSectionGroupId, bibleBookInfo.SectionName);

                var bibleBook = bibleInfo.Books.FirstOrDefault(book => book.Index == bibleBookInfo.Index);
                if (bibleBook == null)
                    throw new Exception("Manifest.xml has Bible books that do not exist in bible.xml");

                foreach (var chapter in bibleBook.Chapters)
                {
                    if (logger != null)
                        logger.LogMessage("{0} '{1} {2}'", BibleCommon.Resources.Constants.ProcessChapter, bibleBookInfo.Name, chapter.Index);

                    GenerateChapterPage(ref oneNoteApp, chapter, bookSectionId, moduleInfo, bibleBookInfo, bibleInfo, currentStrongPrefix);
                }
            }

            OneNoteUtils.UseOneNoteAPI(ref oneNoteApp, (oneNoteAppSafe) =>
            {
                oneNoteAppSafe.SyncHierarchy(SettingsManager.Instance.NotebookId_SupplementalBible);            
            });
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

        private static void UnlockNotebooks(ref Application oneNoteApp, bool unlockBible, bool unlockSupplementalBible, ICustomLogger logger)
        {
            try
            {
                if (unlockBible)                
                    OneNoteLocker.UnlockBible(ref oneNoteApp, true, () => logger.AbortedByUser);
            }
            catch (NotSupportedException)
            {
                //todo: log it
            }

            try
            {
                if (unlockSupplementalBible)
                    OneNoteLocker.UnlockSupplementalBible(ref oneNoteApp, true, () => logger.AbortedByUser);
            }
            catch (NotSupportedException)
            {
                //todo: log it
            }
        }

        /// <summary>
        /// Link first supplemental module with primary Bible
        /// </summary>
        /// <param name="oneNoteApp"></param>
        /// <param name="strongTermLinksCache"></param>
        /// <param name="logger"></param>
        /// <returns></returns>
        public static BibleParallelTranslationConnectionResult LinkSupplementalBibleWithPrimaryBible(ref Application oneNoteApp,
            Dictionary<string, string> strongTermLinksCache, ICustomLogger logger)
        {
            if (string.IsNullOrEmpty(SettingsManager.Instance.GetValidSupplementalBibleNotebookId(ref oneNoteApp, true)) 
                || SettingsManager.Instance.SupplementalBibleModules.Count == 0)
                throw new NotConfiguredException("Supplemental Bible does not exists.");

            XmlNamespaceManager xnm = OneNoteUtils.GetOneNoteXNM();
            var nms = XNamespace.Get(Constants.OneNoteXmlNs);

            string supplementalModuleShortName = SettingsManager.Instance.SupplementalBibleModules.First().ModuleName;                        

            UnlockNotebooks(ref oneNoteApp, true, false, logger);

            var isOneNote2010 = true; // OneNoteUtils.IsOneNote2010Cached(oneNoteApp);

            BibleParallelTranslationConnectionResult result;
            using (var bibleTranslationManager = new BibleParallelTranslationManager(
                            supplementalModuleShortName, SettingsManager.Instance.ModuleShortName,
                            SettingsManager.Instance.NotebookId_SupplementalBible))
            {
                if (bibleTranslationManager.BaseModuleInfo.Type == Common.ModuleType.Strong)
                    if (strongTermLinksCache == null)
                        throw new ArgumentNullException("strongTermLinksCache");                

                bibleTranslationManager.Logger = logger;                               

                var linkResult = new List<Exception>();

                var oneNoteTemp = oneNoteApp;
                result = bibleTranslationManager.IterateBaseBible(
                    (chapterPageDoc, chapterPointer) =>
                    {
                        ApplicationCache.Instance.CommitAllModifiedPages(ref oneNoteTemp, false, pageContent => pageContent.PageType == ApplicationCache.PageType.Bible, null, null);

                        int styleIndex = QuickStyleManager.AddQuickStyleDef(chapterPageDoc, QuickStyleManager.StyleForStrongName, QuickStyleManager.PredefinedStyles.GrayHyperlink, xnm);

                        return new BibleIteratorArgs() { ChapterDocument = chapterPageDoc, StrongStyleIndex = styleIndex };
                    }, true, true,
                    (baseVersePointer, parallelVerse, bibleIteratorArgs) =>
                    {
                        if (!parallelVerse.IsEmpty || parallelVerse.IsPartOfBigVerse || parallelVerse.HasValueEvenIfEmpty)
                        {
                            linkResult.AddRange(
                                LinkPrimaryBibleAndSupplementalVerses(ref oneNoteTemp, baseVersePointer, parallelVerse, bibleIteratorArgs,
                                            bibleTranslationManager.BaseModuleInfo.Type == Common.ModuleType.Strong, strongTermLinksCache,
                                            bibleTranslationManager.BaseModuleInfo.ShortName,
                                            bibleTranslationManager.BaseModuleInfo.BibleStructure.Alphabet, isOneNote2010, xnm, nms));
                        }
                    });
                oneNoteApp = oneNoteTemp;
                oneNoteTemp = null;

                result.Errors.AddRange(linkResult);
            }

            ApplicationCache.Instance.CommitAllModifiedPages(ref oneNoteApp, false, pageContent => pageContent.PageType == ApplicationCache.PageType.Bible, null, null);

            return result;
        }

        public static BibleParallelTranslationConnectionResult AddParallelBible(ref Application oneNoteApp, ModuleInfo module, 
                Dictionary<string, string> strongTermLinksCache, ICustomLogger logger)
        {
            if (string.IsNullOrEmpty(SettingsManager.Instance.GetValidSupplementalBibleNotebookId(ref oneNoteApp, true))
                || SettingsManager.Instance.SupplementalBibleModules.Count == 0)
                throw new NotConfiguredException();

            SettingsManager.Instance.SupplementalBibleModules.Add(new StoredModuleInfo(module.ShortName, module.Version));
            SettingsManager.Instance.Save();

            string oldTestamentName = null;
            int? oldTestamentSectionsCount = null;
            string oldTestamentStrongPrefix = null;
            string newTestamentName = null;
            int? newTestamentSectionsCount = null;
            string newTestamentStrongPrefix = null;            

            BibleParallelTranslationConnectionResult result = null;
            XmlNamespaceManager xnm = OneNoteUtils.GetOneNoteXNM();
            var linkResult = new List<Exception>();
            var isOneNote2010 = true; // OneNoteUtils.IsOneNote2010Cached(oneNoteApp);  // пока ещё окончательно не разобрался с проблемой ссылок в Стронге для OneNote 2013.

            using (var bibleTranslationManager = new BibleParallelTranslationManager(
                SettingsManager.Instance.SupplementalBibleModules.First().ModuleName, module.ShortName,
                SettingsManager.Instance.NotebookId_SupplementalBible))
            {
                UnlockNotebooks(ref oneNoteApp, false, true, logger);

                GetTestamentInfo(bibleTranslationManager.ParallelModuleInfo, ContainerType.OldTestament, out oldTestamentName, out oldTestamentSectionsCount, out oldTestamentStrongPrefix);
                GetTestamentInfo(bibleTranslationManager.ParallelModuleInfo, ContainerType.NewTestament, out newTestamentName, out newTestamentSectionsCount, out newTestamentStrongPrefix);

                var oneNoteTemp = oneNoteApp;
                bibleTranslationManager.Logger = logger;
                result = bibleTranslationManager.IterateBaseBible(
                    (chapterPageDoc, chapterPointer) =>
                    {
                        if (UpdateSupplementalModulesMetadata(ref oneNoteTemp, chapterPageDoc, chapterPointer, module, xnm))
                        {

                            var tableEl = NotebookGenerator.GetPageTable(chapterPageDoc, xnm);
                            var bibleIndex = NotebookGenerator.AddColumnToTable(tableEl, SettingsManager.Instance.PageWidth_Bible, xnm);
                            NotebookGenerator.AddParallelBibleTitle(chapterPageDoc, tableEl,
                                bibleTranslationManager.ParallelModuleInfo.DisplayName, bibleIndex, bibleTranslationManager.ParallelModuleInfo.Locale, xnm);

                            int styleIndex = QuickStyleManager.AddQuickStyleDef(chapterPageDoc, QuickStyleManager.StyleForStrongName, QuickStyleManager.PredefinedStyles.GrayHyperlink, xnm);

                            var strongPrefix = bibleTranslationManager.ParallelModuleInfo.Type == Common.ModuleType.Strong
                                ? GetStrongPrefix(chapterPointer.BookIndex, (oldTestamentSectionsCount ?? newTestamentSectionsCount).Value, oldTestamentStrongPrefix, newTestamentStrongPrefix)
                                : null;

                            return new BibleIteratorArgs() { BibleIndex = bibleIndex, TableElement = tableEl, StrongStyleIndex = styleIndex, StrongPrefix = strongPrefix };
                        }
                        else
                            return new BibleIteratorArgs() { NotNeedToProcessVerses = true, NotNeedToUpdateChapter = true };
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
                                if (isOneNote2010)
                                    QuickStyleManager.SetQuickStyleDefForCell(cell, bibleIteratorArgs.StrongStyleIndex, xnm);
                            }
                        }
                    });

                oneNoteApp = oneNoteTemp;
                oneNoteTemp = null;
            }

            result.Errors.AddRange(linkResult);            

            return result;
        }
        
        private static bool UpdateSupplementalModulesMetadata(ref Application oneNoteApp, XDocument chapterPageDoc, SimpleVersePointer chapterPointer, ModuleInfo module,
            XmlNamespaceManager xnm)
        {
            var supplementalModulesMetadata = OneNoteUtils.GetElementMetaData(chapterPageDoc.Root, Consts.Constants.Key_EmbeddedSupplementalModules, xnm);
            if (string.IsNullOrEmpty(supplementalModulesMetadata))
                throw new InvalidOperationException(string.Format("Chapter page metadata was not found: {0}", chapterPointer));

            var supplementalModulesInfo = EmbeddedModuleInfo.Deserialize(supplementalModulesMetadata);
            if (!supplementalModulesInfo.Any(sm => sm.ModuleName == module.ShortName))
            {
                supplementalModulesInfo.Add(new EmbeddedModuleInfo(module.ShortName, module.Version, supplementalModulesInfo.Count));

                OneNoteUtils.UpdateElementMetaData(chapterPageDoc.Root, Consts.Constants.Key_EmbeddedSupplementalModules,
                    EmbeddedModuleInfo.Serialize(supplementalModulesInfo), xnm);

                return true;
            }
            else
                return false;
        }

        public static void CloseSupplementalBible(ref Application oneNoteApp, bool removeStrongDictionaryFromNotebook, Func<bool> checkIfExternalProcessAborted = null)
        {            
            OneNoteUtils.CloseNotebookSafe(ref oneNoteApp, SettingsManager.Instance.NotebookId_SupplementalBible, checkIfExternalProcessAborted);

            foreach (var parallelModuleName in SettingsManager.Instance.SupplementalBibleModules)
            {                
                var moduleInfo = ModulesManager.GetModuleInfo(parallelModuleName.ModuleName);
                if (moduleInfo.Type == Common.ModuleType.Strong)
                {
                    DictionaryManager.RemoveDictionary(ref oneNoteApp, parallelModuleName.ModuleName, removeStrongDictionaryFromNotebook);
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

        public static RemoveResult RemoveSupplementalBibleModule(ref Application oneNoteApp, string moduleShortName, bool removeStrongDictionaryFromNotebook, ICustomLogger logger)
        {
            if (SettingsManager.Instance.SupplementalBibleModules.Count <= 1)
            {
                CloseSupplementalBible(ref oneNoteApp, removeStrongDictionaryFromNotebook, () => logger.AbortedByUser);
                return RemoveResult.RemoveSupplementalBible;
            }
            else
            {
                var storedModuleInfo = SettingsManager.Instance.SupplementalBibleModules.FirstOrDefault(m => m.ModuleName == moduleShortName);

                if (storedModuleInfo == null)
                    throw new ArgumentException(string.Format("Module '{0}' can not be found in Supplemental Bible", moduleShortName));

                var moduleInfo = ModulesManager.GetModuleInfo(moduleShortName);
                if (moduleInfo.Type == Common.ModuleType.Strong)
                    DictionaryManager.RemoveDictionary(ref oneNoteApp, moduleShortName, removeStrongDictionaryFromNotebook);

                SettingsManager.Instance.SupplementalBibleModules.Remove(storedModuleInfo);
                SettingsManager.Instance.Save();               

                using (var bibleTranslationManager = new BibleParallelTranslationManager(
                   SettingsManager.Instance.SupplementalBibleModules.First().ModuleName, moduleShortName,
                   SettingsManager.Instance.NotebookId_SupplementalBible))
                {
                    UnlockNotebooks(ref oneNoteApp, false, true, logger);

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
            var styleEl = pageDoc.Root.XPathSelectElement(string.Format("one:QuickStyleDef[@name=\"{0}\"]", QuickStyleManager.StyleForStrongName), xnm);
            if (styleEl != null)  // значит видимо есть Библия Стронга на текущей странице
            {
                string searchTemplate = "</a>";
                XNamespace nms = XNamespace.Get(Constants.OneNoteXmlNs);

                foreach (var textEl in pageDoc.Root.XPathSelectElements(string.Format("//one:OE[@quickStyleIndex=\"{0}\"]/one:T", (string)styleEl.Attribute("index")), xnm))
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

        private static List<Exception> LinkPrimaryBibleAndSupplementalVerses(ref Application oneNoteApp, SimpleVersePointer baseVersePointer,
            SimpleVerse parallelVerse, BibleIteratorArgs bibleIteratorArgs, bool isStrong, Dictionary<string, string> strongTermLinksCache, 
            string strongModuleShortName, string alphabet, bool isOneNote2010, XmlNamespaceManager xnm, XNamespace nms)
        {
            var result = new List<Exception>();            

            var parallelVersePointer = parallelVerse.ToVersePointer(SettingsManager.Instance.CurrentModuleCached);
            var primaryBibleObjectsSearchResult = HierarchySearchManager.GetHierarchyObject(ref oneNoteApp,
                    SettingsManager.Instance.NotebookId_Bible, ref parallelVersePointer, HierarchySearchManager.FindVerseLevel.AllVerses, null, null);

            if (primaryBibleObjectsSearchResult.ResultType != BibleHierarchySearchResultType.Successfully
                || primaryBibleObjectsSearchResult.HierarchyStage != BibleHierarchyStage.ContentPlaceholder)
                throw new VerseNotFoundException(parallelVerse, SettingsManager.Instance.ModuleShortName, BaseVersePointerException.Severity.Error);

            VerseNumber? baseVerseNumber;
            string verseTextWithoutNumber;
            var baseVerseEl = OneNoteUtils.NormalizeTextElement(
                                    HierarchySearchManager.FindVerse(bibleIteratorArgs.ChapterDocument, false, baseVersePointer.Verse, xnm,
                                    out baseVerseNumber, out verseTextWithoutNumber));            
                
            var baseChapterPageId = (string)bibleIteratorArgs.ChapterDocument.Root.Attribute("ID");
            var baseVerseElementId = (string)baseVerseEl.Parent.Attribute("objectID");            

            if (!parallelVerse.IsEmpty)
            LinkMainBibleVersesToSupplementalBibleVerse(ref oneNoteApp, baseChapterPageId, baseVerseElementId, parallelVerse, primaryBibleObjectsSearchResult, xnm, nms);

            LinkSupplementalBibleVerseToMainBibleVerseAndToStrongDictionary(ref oneNoteApp, baseVersePointer, baseVerseEl, baseVerseNumber, parallelVersePointer, verseTextWithoutNumber, primaryBibleObjectsSearchResult, 
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
                                                                    out temp, out htmlBreakIndex, null, StringSearchMode.SearchNumber);
                if (!string.IsNullOrEmpty(strongNumber))
                {
                    string prefix = StringUtils.GetPrevString(verseText, cursorPosition, new SearchMissInfo(null, SearchMissInfo.MissMode.CancelOnMissFound), alphabet,
                                                                    out temp, out temp, null, StringSearchMode.SearchFirstChar);
                    if (!string.IsNullOrEmpty(prefix) && prefix.Length == 1 && StringUtils.IsCharAlphabetical(prefix[0], alphabet))
                    {
                        string strongTerm = string.Format("{0}{1:0000}", prefix, int.Parse(strongNumber));
                        if (strongTermLinksCache.ContainsKey(strongTerm))
                        {
                            var termLink = new DictionaryTermLink(strongTermLinksCache[strongTerm]).Href;
                            string link = string.Format("<a href=\"{0}\"><span style='vertical-align:super;'>{1}</span></a>",
                                SettingsManager.Instance.UseProxyLinksForStrong || string.IsNullOrEmpty(termLink)
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

        private static void LinkSupplementalBibleVerseToMainBibleVerseAndToStrongDictionary(ref Application oneNoteApp, 
            SimpleVersePointer baseVersePointer, XElement baseVerseEl, VerseNumber? baseVerseNumber, VersePointer parallelVersePointer,
            string verseTextWithoutNumber, BibleSearchResult primaryBibleObjectsSearchResult,
            bool isStrong, int strongStyleIndex, Dictionary<string, string> strongTermLinksCache, string strongModuleShortName, string alphabet, bool isOneNote2010,
            ref List<Exception> result, XNamespace nms)
        {
            if (baseVersePointer.VerseNumber != baseVerseNumber)            
                result.Add(
                    new InvalidOperationException(
                        string.Format("baseVerseNumber != baseVersePointer (baseVerseNumber = '{0}', baseVersePointer = '{1}')", baseVerseNumber, baseVersePointer)));

            string linkToParallelVerse = OneNoteUtils.GetOrGenerateLink(ref oneNoteApp, baseVerseNumber.ToString(),
                                            SettingsManager.Instance.UseProxyLinksForBibleVerses 
                                                ? OpenBibleVerseHandler.GetCommandUrlStatic(parallelVersePointer, SettingsManager.Instance.ModuleShortName) 
                                                : primaryBibleObjectsSearchResult.HierarchyObjectInfo.VerseInfo.ProxyHref,
                                            primaryBibleObjectsSearchResult.HierarchyObjectInfo.PageId, primaryBibleObjectsSearchResult.HierarchyObjectInfo.VerseContentObjectId, 
                                            SettingsManager.Instance.UseProxyLinksForBibleVerses 
                                                ? null
                                                : Consts.Constants.QueryParameter_BibleVerse);

            string versePart = verseTextWithoutNumber;

            if (isStrong)
            {
                if (isOneNote2010)
                    baseVerseEl.Parent.SetAttributeValue("quickStyleIndex", strongStyleIndex);
                versePart = ProcessStrongVerse(versePart, strongTermLinksCache, strongModuleShortName, alphabet, isOneNote2010, ref result);
            }

            baseVerseEl.Value = string.Format("{0}<span> </span>{1}", linkToParallelVerse, versePart);
        }

        private static void LinkMainBibleVersesToSupplementalBibleVerse(ref Application oneNoteApp, string baseChapterPageId, string baseVerseElementId,
            SimpleVerse parallelVerse, BibleSearchResult baseBibleObjectsSearchResult, XmlNamespaceManager xnm, XNamespace nms)
        {           
            if (parallelVerse.PartIndex.GetValueOrDefault(0) == 0 && !parallelVerse.IsEmpty 
                //&& !string.IsNullOrEmpty(parallelVerse.VerseContent)  интересно зачем такое условие сделал? Ведь одно дело он IsEmpty, а другое дело просто пустой...
                )  // если PartIndex > 0, значит этот стих мы уже привязали
            {
                var parallelChapterPageDoc = PrepareMainBibleTable(ref oneNoteApp, baseBibleObjectsSearchResult.HierarchyObjectInfo.PageId);

                string linkToBaseVerse = string.Format("<font size='2pt'>{0}</font>",
                                            OneNoteUtils.GenerateLink(ref oneNoteApp, SettingsManager.Instance.SupplementalBibleLinkName, baseChapterPageId, baseVerseElementId));

                foreach (var parallelVerseElementId in baseBibleObjectsSearchResult.HierarchyObjectInfo.GetAllObjectsIds())
                {                    
                    var bibleCell = parallelChapterPageDoc.Content.Root
                                    .XPathSelectElement(string.Format("//one:OE[@objectID=\"{0}\"]", parallelVerseElementId.ObjectId), xnm).Parent.Parent;
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

        private static ApplicationCache.PageContent PrepareMainBibleTable(ref Application oneNoteApp, string mainBibleChapterPageId)
        {
            var parallelChapterPageDoc = ApplicationCache.Instance.GetPageContent(ref oneNoteApp, mainBibleChapterPageId, ApplicationCache.PageType.Bible);
            var parallelBibleTableElement = NotebookGenerator.GetPageTable(parallelChapterPageDoc.Content, parallelChapterPageDoc.Xnm);

            var columnsCount = parallelBibleTableElement.XPathSelectElements("one:Columns/one:Column", parallelChapterPageDoc.Xnm).Count();
            if (columnsCount == 2)
                NotebookGenerator.AddColumnToTable(parallelBibleTableElement, NotebookGenerator.MinimalCellWidth, parallelChapterPageDoc.Xnm);
            parallelChapterPageDoc.WasModified = true;

            return parallelChapterPageDoc;
        }


        private static void GenerateChapterPage(ref Application oneNoteApp, CHAPTER chapter, string bookSectionId,
           ModuleInfo moduleInfo, BibleBookInfo bibleBookInfo, XMLBIBLE bibleInfo, string strongPrefix)
        {
            string chapterPageName = string.Format(!string.IsNullOrEmpty(bibleBookInfo.ChapterPageNameTemplate) 
                                                        ? bibleBookInfo.ChapterPageNameTemplate
                                                        : moduleInfo.BibleStructure.ChapterPageNameTemplate, 
                                                   chapter.Index, bibleBookInfo.Name);

            XmlNamespaceManager xnm;
            var currentChapterDoc = NotebookGenerator.AddPage(ref oneNoteApp, bookSectionId, chapterPageName, 1, moduleInfo.Locale, out xnm);

            OneNoteUtils.UpdateElementMetaData(currentChapterDoc.Root, Consts.Constants.Key_EmbeddedSupplementalModules,
                EmbeddedModuleInfo.Serialize(new List<EmbeddedModuleInfo>() { new EmbeddedModuleInfo(moduleInfo.ShortName, moduleInfo.Version, 0) }), xnm);

            var currentTableElement = NotebookGenerator.AddTableToPage(currentChapterDoc, false, xnm, new CellInfo(SettingsManager.Instance.PageWidth_Bible));

            NotebookGenerator.AddParallelBibleTitle(currentChapterDoc, currentTableElement, moduleInfo.DisplayName, 0, moduleInfo.Locale, xnm);

            foreach (var verse in chapter.Verses)
            {                
                NotebookGenerator.AddVerseRowToTable(currentTableElement, BIBLEBOOK.GetFullVerseString(verse.Index, verse.TopIndex, verse.GetValue(true, strongPrefix)), 0, moduleInfo.Locale);
            }

            OneNoteUtils.UpdatePageContentSafe(ref oneNoteApp, currentChapterDoc, xnm);            
        }

        private static string GetCurrentSectionGroupId(ref Application oneNoteApp, string currentSectionGroupId, 
            string oldTestamentName, int? oldTestamentSectionsCount, string newTestamentName, int? newTestamentSectionsCount, int i)
        {
            if (string.IsNullOrEmpty(currentSectionGroupId))
            {
                currentSectionGroupId
                    = (string)NotebookGenerator.AddRootSectionGroupToNotebook(ref oneNoteApp,
                        SettingsManager.Instance.NotebookId_SupplementalBible, oldTestamentName ?? newTestamentName).Attribute("ID");                
            }
            else if (i == (oldTestamentSectionsCount ?? newTestamentSectionsCount))  // если только один завет в модуле, то до сюда и не должен дойти
            {
                currentSectionGroupId
                    = (string)NotebookGenerator.AddRootSectionGroupToNotebook(ref oneNoteApp,
                        SettingsManager.Instance.NotebookId_SupplementalBible, newTestamentName).Attribute("ID");             
            }            

            return currentSectionGroupId;
        }             
    }
}
