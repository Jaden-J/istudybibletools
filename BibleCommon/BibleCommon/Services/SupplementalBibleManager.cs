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

namespace BibleCommon.Services
{
    public static class SupplementalBibleManager
    {
        public static void CreateSupplementalBible(Application oneNoteApp, string moduleShortName)
        {
            if (!string.IsNullOrEmpty(SettingsManager.Instance.NotebookId_SupplementalBible))            
                if (!OneNoteUtils.NotebookExists(oneNoteApp, SettingsManager.Instance.NotebookId_SupplementalBible))
                    SettingsManager.Instance.NotebookId_SupplementalBible = null;
            

            if (string.IsNullOrEmpty(SettingsManager.Instance.NotebookId_SupplementalBible))            
                SettingsManager.Instance.NotebookId_SupplementalBible = NotebookGenerator.CreateNotebook(oneNoteApp, Resources.Constants.SupplementalBibleName);                            
            else
                throw new InvalidOperationException("Supplemental Bible already exists");

            SettingsManager.Instance.SupplementalBibleModules.Clear();
            SettingsManager.Instance.SupplementalBibleModules.Add(moduleShortName);
            SettingsManager.Instance.Save();
            
            string currentSectionGroupId = null;
            var moduleInfo = ModulesManager.GetModuleInfo(moduleShortName);
            var bibleInfo = ModulesManager.GetModuleBibleInfo(moduleShortName);            

            for (int i = 0; i < moduleInfo.BibleStructure.BibleBooks.Count; i++)
            {
                var bibleBookInfo = moduleInfo.BibleStructure.BibleBooks[i];
                bibleBookInfo.SectionName = NotebookGenerator.GetBibleBookSectionName(bibleBookInfo.Name, i, moduleInfo.BibleStructure.OldTestamentBooksCount);

                currentSectionGroupId = GetCurrentSectionGroupId(oneNoteApp, currentSectionGroupId, moduleInfo, i);                

                var bookSectionId = NotebookGenerator.AddBookSectionToBibleNotebook(oneNoteApp, currentSectionGroupId, bibleBookInfo.SectionName, bibleBookInfo.Name);

                var bibleBook = bibleInfo.Content.Books.FirstOrDefault(book => book.Index == bibleBookInfo.Index);
                if (bibleBook == null)
                    throw new Exception("Manifest.xml has Bible books that do not exist in bible.xml");

                foreach (var chapter in bibleBook.Chapters)
                {   
                    GenerateChapterPage(oneNoteApp, chapter, bookSectionId, moduleInfo, bibleBookInfo, bibleInfo);                    
                }                
            }         
        }

        public static BibleParallelTranslationConnectionResult LinkSupplementalBibleWithMainBible(Application oneNoteApp, int supplementalModuleIndex)
        {
            if (supplementalModuleIndex != 0)
                throw new NotSupportedException("supplementalModuleIndex != 0");

            string supplementalModuleSortName = SettingsManager.Instance.SupplementalBibleModules[supplementalModuleIndex];
            bool needToLinkMainBibleToSupplementalBible = supplementalModuleIndex == 0;

            XmlNamespaceManager xnm = OneNoteUtils.GetOneNoteXNM();
            var nms = XNamespace.Get(Constants.OneNoteXmlNs);

            var bibleTranslationManager = new BibleParallelTranslationManager(oneNoteApp,
                            supplementalModuleSortName, SettingsManager.Instance.ModuleName,
                            SettingsManager.Instance.NotebookId_SupplementalBible);

            var result = bibleTranslationManager.IterateBaseBible(chapterPageDoc =>
                {
                    return new BibleIteratorArgs() { ChapterDocument = chapterPageDoc };
                }, true,
                (baseVersePointer, parallelVerse, bibleIteratorArgs) =>
                {
                    LinkdMainBibleAndSupplementalVerses(oneNoteApp, baseVersePointer, parallelVerse, bibleIteratorArgs, xnm, nms);
                });

            OneNoteProxy.Instance.CommitAllModifiedPages(oneNoteApp, pageContent => pageContent.PageType == OneNoteProxy.PageType.Bible, null, null);

            return result;
        }

        private static void LinkdMainBibleAndSupplementalVerses(Application oneNoteApp, SimpleVersePointer baseVersePointer,
            SimpleVerse parallelVerse, BibleIteratorArgs bibleIteratorArgs, XmlNamespaceManager xnm, XNamespace nms)
        {

            var baseBibleObjectsSearchResult = HierarchySearchManager.GetHierarchyObject(oneNoteApp,
                    SettingsManager.Instance.NotebookId_Bible, parallelVerse.ToVersePointer(SettingsManager.Instance.CurrentModule), true);

            if (baseBibleObjectsSearchResult.ResultType != HierarchySearchManager.HierarchySearchResultType.Successfully
                && baseBibleObjectsSearchResult.HierarchyStage != HierarchySearchManager.HierarchyStage.ContentPlaceholder)
                throw new VerseNotFoundException(parallelVerse, BaseVersePointerException.Severity.Error);

            var baseVerseEl = HierarchySearchManager.FindVerse(bibleIteratorArgs.ChapterDocument, false, baseVersePointer.Verse, xnm);
            var baseChapterPageId = (string)bibleIteratorArgs.ChapterDocument.Root.Attribute("ID").Value;
            var baseVerseElementId = (string)baseVerseEl.Attribute("objectID").Value;

            LinkMainBibleVersesToSupplementalBibleVerse(oneNoteApp, baseChapterPageId, baseVerseElementId, parallelVerse, baseBibleObjectsSearchResult, nms);
            LinkSupplementalBibleVerseToMainBibleVerse(oneNoteApp, baseVersePointer, baseVerseEl, baseBibleObjectsSearchResult);           
        }

        private static void LinkSupplementalBibleVerseToMainBibleVerse(Application oneNoteApp, SimpleVersePointer baseVersePointer, XElement baseVerseEl, HierarchySearchManager.HierarchySearchResult baseBibleObjectsSearchResult)
        {
            int textBreakIndex, htmlBreakIndex;
            var baseVerseNumber = StringUtils.GetNextString(baseVerseEl.Value, 0, new SearchMissInfo(0, SearchMissInfo.MissMode.CancelOnMissFound), out textBreakIndex, out htmlBreakIndex);

            if (baseVerseNumber != baseVersePointer.Verse.ToString())
                throw new InvalidOperationException(
                    string.Format("baseVerseNumber != baseVersePointer (baseVerseNumber = '{0}', baseVersePointer = '{1}')", baseVerseNumber, baseVersePointer));

            string linkToParallelVerse = OneNoteUtils.GenerateHref(oneNoteApp, baseVerseNumber,
                baseBibleObjectsSearchResult.HierarchyObjectInfo.PageId, baseBibleObjectsSearchResult.HierarchyObjectInfo.ContentObjectId);

            baseVerseEl.Value = string.Format("{0} {1}", linkToParallelVerse, baseVerseEl.Value.Substring(htmlBreakIndex) + 1);
        }

        private static void LinkMainBibleVersesToSupplementalBibleVerse(Application oneNoteApp, string baseChapterPageId, string baseVerseElementId, 
            SimpleVerse parallelVerse, HierarchySearchManager.HierarchySearchResult baseBibleObjectsSearchResult, XNamespace nms)
        {           
            if (parallelVerse.PartIndex.GetValueOrDefault(0) == 0)  // если PartIndex > 0, значит этот стих мы уже привязали
            {
                var parallelChapterPageDoc = PrepareMainBibleTable(oneNoteApp, baseBibleObjectsSearchResult.HierarchyObjectInfo.PageId);

                string linkToBaseVerse = OneNoteUtils.GenerateHref(oneNoteApp, SettingsManager.Instance.SupplementalBibleLinkName, baseChapterPageId, baseVerseElementId);

                foreach (var parallelVerseElementId in baseBibleObjectsSearchResult.HierarchyObjectInfo.GetAllObjectsIds())
                {
                    var cell = parallelChapterPageDoc.Content.Root.XPathSelectElement(string.Format("//one:Cell[@objectID='{0}']", parallelVerseElementId));
                    var row = cell.Parent;
                    if (row.Elements().Count() == 3)
                        row.Elements().Last().Value = linkToBaseVerse;
                    else
                        row.Add(NotebookGenerator.GetCell(linkToBaseVerse, string.Empty, nms));
                }
            }
        }

        private static OneNoteProxy.PageContent PrepareMainBibleTable(Application oneNoteApp, string mainBibleChapterPageId)
        {
            var parallelChapterPageDoc = OneNoteProxy.Instance.GetPageContent(oneNoteApp, mainBibleChapterPageId, OneNoteProxy.PageType.Bible);
            var parallelBibleTableElement = NotebookGenerator.GetBibleTable(parallelChapterPageDoc.Content, parallelChapterPageDoc.Xnm);

            var columnsCount = parallelBibleTableElement.XPathSelectElements("one:Columns/one:Column", parallelChapterPageDoc.Xnm).Count();
            if (columnsCount == 2)
                NotebookGenerator.AddColumnToTable(parallelBibleTableElement, NotebookGenerator.MinimalCellWidth, parallelChapterPageDoc.Xnm);
            parallelChapterPageDoc.WasModified = true;

            return parallelChapterPageDoc;
        }               

        //private static void UpdateBibleChapterLinksToSupplementalBible(Application oneNoteApp, string chapterPageId, int chapterIndex, BibleBookInfo bibleBookInfo)
        //{
        //    XmlNamespaceManager xnm;
        //    var nms = XNamespace.Get(Constants.OneNoteXmlNs);           
            
        //    var chapterPageDoc = OneNoteUtils.GetPageContent(oneNoteApp, chapterPageId, out xnm);
        //    var bibleTable = NotebookGenerator.GetBibleTable(chapterPageDoc, xnm);
        //    SimpleVersePointer prevParallelVersePointer = null;

        //    foreach(var cell in bibleTable.XPathSelectElements("one:Row/one:Cell[1]/one:OEChildren/one:OE/one:T", xnm).Skip(1))
        //    {
        //        var verseIndex = StringUtils.GetStringFirstNumber(cell.Value);
        //        if (verseIndex.HasValue)
        //        {
        //            var cellId = (string)cell.Parent.Attribute("objectID").Value;
        //            string link = OneNoteUtils.GenerateHref(oneNoteApp, SettingsManager.Instance.SupplementalBibleLinkName, chapterPageId, cellId);

        //            var baseVersePointer = new SimpleVersePointer(bibleBookInfo.Index, chapterIndex, verseIndex.Value);
        //            var parallelVersePointer = BibleParallelTranslationConnectorManager.GetParallelVersePointer(baseVersePointer,
        //                SettingsManager.Instance.SupplementalBibleModules.First(), SettingsManager.Instance.ModuleName);

        //            if (parallelVersePointer != prevParallelVersePointer)
        //            {
        //                prevParallelVersePointer = parallelVersePointer;
        //                if (!parallelVersePointer.IsEmpty)
        //                {
        //                    var mainBibleChapterDoc = PrepareMainBibleTable(oneNoteApp, parallelVersePointer, out xnm);

        //                    if (mainBibleChapterDoc != null)
        //                    {
        //                        var mainBibleVerseEl = HierarchySearchManager.FindVerse(mainBibleChapterDoc, false, parallelVersePointer.Verse, xnm);
        //                        if (mainBibleVerseEl == null)
        //                            throw new Exception(string.Format("Can not find Bible verse cell for '{0} {1}:{2}'", bibleBookInfo.Name, chapterIndex, verseIndex));

        //                        var mainBibleVerseRowEl = mainBibleVerseEl.Parent.Parent.Parent.Parent;
        //                        var sbCell = mainBibleVerseRowEl.XPathSelectElement("one:Cell[3]/one:OEChildren/one:OE/one:T", xnm);

        //                        if (sbCell == null)
        //                            mainBibleVerseRowEl.Add(NotebookGenerator.GetCell(link, string.Empty, nms));
        //                        else
        //                            sbCell.Value = link;
        //                    }
        //                }
        //            }
        //            else
        //            {
        //                int i = 0;
        //            }
        //        }
        //    }            
        //}    

        private static void GenerateChapterPage(Application oneNoteApp, BibleChapterContent chapter, string bookSectionId,
           ModuleInfo moduleInfo, BibleBookInfo bibleBookInfo, ModuleBibleInfo bibleInfo)
        {
            string chapterPageName = string.Format(moduleInfo.BibleStructure.ChapterSectionNameTemplate, chapter.Index, bibleBookInfo.Name);

            XmlNamespaceManager xnm;
            var currentChapterDoc = NotebookGenerator.AddChapterPageToBibleNotebook(oneNoteApp, bookSectionId, chapterPageName, 1, bibleInfo.Content.Locale, out xnm);

            var currentTableElement = NotebookGenerator.AddTableToBibleChapterPage(currentChapterDoc, SettingsManager.Instance.PageWidth_Bible, xnm);

            NotebookGenerator.AddParallelBibleTitle(currentTableElement, moduleInfo.Name, 0, bibleInfo.Content.Locale, xnm);

            foreach (var verse in chapter.Verses)
            {
                NotebookGenerator.AddVerseRowToBibleTable(currentTableElement, string.Format("{0} {1}", verse.Index, verse.Value), bibleInfo.Content.Locale);
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

        public static BibleParallelTranslationConnectionResult AddParallelBible(Application oneNoteApp, string moduleShortName)
        {
            if (string.IsNullOrEmpty(SettingsManager.Instance.NotebookId_SupplementalBible) || SettingsManager.Instance.SupplementalBibleModules.Count == 0)            
                throw new Exception(BibleCommon.Resources.Constants.Error_SystemIsNotConfigures);

            var bibleTranslationManager = new BibleParallelTranslationManager(oneNoteApp, 
                SettingsManager.Instance.SupplementalBibleModules.First(), moduleShortName, 
                SettingsManager.Instance.NotebookId_SupplementalBible);

            var result = bibleTranslationManager.AddParallelTranslation();            

            SettingsManager.Instance.SupplementalBibleModules.Add(moduleShortName);


            // ещё надо объединить сокращения книг

            SettingsManager.Instance.Save();

            return result;
        }
    }
}
