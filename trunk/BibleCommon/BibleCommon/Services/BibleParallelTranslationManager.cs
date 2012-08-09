using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using BibleCommon.Common;
using Microsoft.Office.Interop.OneNote;
using BibleCommon.Helpers;
using System.Xml;
using BibleCommon.Consts;

namespace BibleCommon.Services
{
    public static class BibleParallelTranslationManager
    {
        private const string supportedModuleMinVersion = "2.0";

        private enum BibleGeneratorDecision
        {
            SameLocation,

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="moduleShortName">Module Directory Name</param>
        /// <param name="translationIndex">0 - основная Библия</param>
        public static void AddParallelTranslation(Application oneNoteApp, string moduleShortName)
        {
            if (!SettingsManager.Instance.CurrentModuleIsCorrect())
                throw new InvalidOperationException("Current module is not correct.");

            if (SettingsManager.Instance.CurrentModule.Version.CompareTo(supportedModuleMinVersion) < 0)
                throw new NotSupportedException(string.Format("Version of current module is {0}.", SettingsManager.Instance.CurrentModule.Version));            
            
            var baseModuleInfo = SettingsManager.Instance.CurrentModule;
            var parallelModuleInfo = ModulesManager.GetModuleInfo(moduleShortName);
            ModulesManager.CheckModule(parallelModuleInfo);  // если с модулем то-то не так, то выдаст ошибку
            if (parallelModuleInfo.Version.CompareTo(supportedModuleMinVersion) < 0)
                throw new NotSupportedException(string.Format("Version of parallel module is {0}.", SettingsManager.Instance.CurrentModule.Version));


            var translationIndex = SettingsManager.Instance.ParallelModules.Count;
            var baseBibleInfo = ModulesManager.GetModuleBibleInfo(baseModuleInfo.ShortName);
            var parallelBibleInfo = ModulesManager.GetModuleBibleInfo(moduleShortName);

            var bibleVersePointersComparisonTable = BibleParallelTranslationConnectorManager.ConnectBibleTranslations(
                                                            baseModuleInfo.BibleTranslationDifferences,
                                                            parallelModuleInfo.BibleTranslationDifferences);                        

            GenerateParallelBibleTables(oneNoteApp, SettingsManager.Instance.NotebookId_Bible,
                baseModuleInfo, baseBibleInfo, parallelModuleInfo, parallelBibleInfo, bibleVersePointersComparisonTable);            
        }

        private static void GenerateParallelBibleTables(Application oneNoteApp, string bibleNotebookId, 
            ModuleInfo baseModuleInfo, ModuleBibleInfo baseBibleInfo, 
            ModuleInfo parallelModuleInfo, ModuleBibleInfo parallelBibleInfo,
            Dictionary<int, SimpleVersePointersComparisonTable> bibleVersePointersComparisonTable)
        {
            foreach (var baseBibleBook in baseBibleInfo.Content.Books)
            {
                var baseBookInfo = baseModuleInfo.BibleStructure.BibleBooks.FirstOrDefault(b => b.Index == baseBibleBook.Index);
                if (baseBookInfo == null)
                    throw new InvalidModuleException(string.Format("Book with index {0} is not found in module manifest", baseBibleBook.Index));

                var parallelBibleBook = parallelBibleInfo.Content.Books.FirstOrDefault(b => b.Index == baseBibleBook.Index);
                if (parallelBibleBook != null)
                {
                    XElement sectionEl = HierarchySearchManager.FindBibleBookSection(oneNoteApp, bibleNotebookId, baseBookInfo.SectionName);
                    if (sectionEl == null)
                        throw new Exception(string.Format("Section with name {0} is not found", baseBookInfo.SectionName));

                    SimpleVersePointersComparisonTable bookVersePointersComparisonTable = bibleVersePointersComparisonTable.ContainsKey(baseBibleBook.Index)
                        ? bibleVersePointersComparisonTable[baseBibleBook.Index] : new SimpleVersePointersComparisonTable();        

                    ProcessBibleBook(oneNoteApp, sectionEl, baseBookInfo, baseBibleBook, baseBookInfo, 
                        parallelBibleBook, bookVersePointersComparisonTable, parallelBibleInfo.Content.Locale);                    
                }
            }
        }

        private static void ProcessBibleBook(Application oneNoteApp, XElement sectionEl, BibleBookInfo baseBibleInfo, 
            BibleBookContent baseBibleBook, BibleBookInfo baseBookInfo, BibleBookContent parallelBibleBook,
            SimpleVersePointersComparisonTable bookVersePointersComparisonTable, string locale)
        {
            XmlNamespaceManager xnm;
            string sectionId = (string)sectionEl.Attribute("ID");
            string sectionName = (string)sectionEl.Attribute("name");

            var sectionPagesEl = OneNoteUtils.GetHierarchyElement(oneNoteApp, sectionId, HierarchyScope.hsPages, out xnm);                        

            foreach (var baseChapter in baseBibleBook.Chapters)
            {
                var chapterPageEl = HierarchySearchManager.FindChapterPage(oneNoteApp, sectionPagesEl.Root, baseChapter.Index, xnm);

                if (chapterPageEl == null)
                    throw new Exception(string.Format("The page for chapter {0} of book {1} does not found", baseChapter.Index, baseBookInfo.Name));

                string chapterPageId = (string)chapterPageEl.Attribute("ID");
                var chapterPageDoc = OneNoteUtils.GetPageContent(oneNoteApp, chapterPageId, out xnm);

                var tableEl = NotebookGenerator.GetBibleTable(chapterPageDoc, xnm);
                int bibleIndex = NotebookGenerator.ExtendBibleTableForParallelTranslation(tableEl, SettingsManager.Instance.PageWidth_Bible, xnm);

                int lastProcessedVerse = 0;
                int lastProcessedChapter = 0;

                foreach (var baseVerse in baseChapter.Verses)
                {
                    var baseVersePointer = new SimpleVersePointer(baseBibleBook.Index, baseChapter.Index, baseVerse.Index);
                    
                    var parallelVerse = GetParallelVerse(baseVersePointer, parallelBibleBook, bookVersePointersComparisonTable, 
                                                                                                lastProcessedChapter, lastProcessedVerse);

                    NotebookGenerator.AddParallelVerseRowToBibleTable(tableEl, parallelVerse.VerseContent, bibleIndex, locale, xnm);                    

                    lastProcessedChapter = parallelVerse.Chapter;
                    lastProcessedVerse = parallelVerse.Verse;
                }                

                oneNoteApp.UpdatePageContent(chapterPageDoc.ToString(), DateTime.MinValue, Constants.CurrentOneNoteSchema);
            }            
        }        

        private static SimpleVerse GetParallelVerse(SimpleVersePointer baseVersePointer, BibleBookContent parallelBibleBook,
            SimpleVersePointersComparisonTable bookVersePointersComparisonTable,  int lastProcessedChapter, int lastProcessedVerse)
        {
            SimpleVersePointer parallelVersePointer = bookVersePointersComparisonTable.ContainsKey(baseVersePointer) 
                                                    ? bookVersePointersComparisonTable[baseVersePointer]
                                                    : baseVersePointer;


            if (lastProcessedChapter > 0 && parallelVersePointer.Chapter > lastProcessedChapter)
            {
                if (parallelBibleBook.Chapters[lastProcessedChapter - 1].Verses.Count > lastProcessedVerse)
                {
                    parallelVersePointer = new SimpleVersePointer(baseVersePointer.BookIndex, lastProcessedChapter, lastProcessedVerse + 1);
                }
            }
            else
            {
                if (lastProcessedVerse > 0 && parallelVersePointer.Verse > lastProcessedVerse + 1)
                {
                    parallelVersePointer = new SimpleVersePointer(baseVersePointer.BookIndex, lastProcessedChapter, lastProcessedChapter + 1);
                }
            }

            return GetParallelVerse(baseVersePointer, parallelVersePointer, parallelBibleBook);           
        }

        private static SimpleVerse GetParallelVerse(SimpleVersePointer baseVersePointer, SimpleVersePointer parallelVersePointer, BibleBookContent parallelBibleBook)
        {
            string verseContent = parallelBibleBook.Chapters[parallelVersePointer.Chapter - 1].Verses[parallelVersePointer.Verse - 1].Value;

            if (parallelVersePointer.Chapter != baseVersePointer.Chapter)
                verseContent = string.Format("{0}:{1}", parallelVersePointer.Chapter, verseContent);

            return new SimpleVerse(parallelVersePointer, verseContent);
        }

        private static int? GetChapter(string pageName, string bookName)
        {
            int? result = null;

            if (!string.IsNullOrEmpty(pageName) && !string.IsNullOrEmpty(bookName))
            {
                if (StringUtils.IsDigit(bookName[0]))  // то есть имя книги начинается с цифры (2Петра)
                    result = StringUtils.GetStringFirstNumber(pageName.Substring(1));
                else
                    result = StringUtils.GetStringFirstNumber(pageName);
            }

            return result;         
        }

        public static void RemoveLastParallelTranslation(Application oneNoteApp)
        {
            throw new NotImplementedException();
        }
    }
}
