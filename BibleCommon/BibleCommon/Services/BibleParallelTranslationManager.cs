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
    public class BibleParallelTranslationConnectionResult
    {
        public List<BaseVersePointerException> Errors { get; set; }

        public BibleParallelTranslationConnectionResult()
        {
            this.Errors = new List<BaseVersePointerException>();
        }
    }

    public static class BibleParallelTranslationManager
    {
        private const string supportedModuleMinVersion = "2.0";        

        /// <summary>
        /// 
        /// </summary>
        /// <param name="moduleShortName">Module Directory Name</param>
        /// <param name="translationIndex">0 - основная Библия</param>
        public static BibleParallelTranslationConnectionResult AddParallelTranslation(Application oneNoteApp, string moduleShortName)
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

            return GenerateParallelBibleTables(oneNoteApp, SettingsManager.Instance.NotebookId_Bible,
                baseModuleInfo, baseBibleInfo, parallelModuleInfo, parallelBibleInfo, bibleVersePointersComparisonTable);            
        }

        private static BibleParallelTranslationConnectionResult GenerateParallelBibleTables(Application oneNoteApp, string bibleNotebookId, 
            ModuleInfo baseModuleInfo, ModuleBibleInfo baseBibleInfo, 
            ModuleInfo parallelModuleInfo, ModuleBibleInfo parallelBibleInfo,
            Dictionary<int, SimpleVersePointersComparisonTable> bibleVersePointersComparisonTable)
        {
            var result = new BibleParallelTranslationConnectionResult();

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

                    result.Errors.AddRange(ProcessBibleBook(oneNoteApp, sectionEl, baseBookInfo, baseBibleBook, baseBookInfo, 
                        parallelBibleBook, parallelModuleInfo.BibleTranslationDifferences.PartVersesAlphabet, 
                        bookVersePointersComparisonTable, parallelBibleInfo.Content.Locale));                    
                }
            }

            return result;
        }

        private static List<BaseVersePointerException> ProcessBibleBook(Application oneNoteApp, XElement sectionEl, BibleBookInfo baseBibleInfo, 
            BibleBookContent baseBibleBook, BibleBookInfo baseBookInfo, BibleBookContent parallelBibleBook, string partVersesAlphabet,
            SimpleVersePointersComparisonTable bookVersePointersComparisonTable, string locale)
        {
            var result = new List<BaseVersePointerException>();

            XmlNamespaceManager xnm;
            string sectionId = (string)sectionEl.Attribute("ID");
            string sectionName = (string)sectionEl.Attribute("name");

            var sectionPagesEl = OneNoteUtils.GetHierarchyElement(oneNoteApp, sectionId, HierarchyScope.hsPages, out xnm);

            int lastProcessedChapter = 0;
            int lastProcessedVerse = 0;     

            foreach (var baseChapter in baseBibleBook.Chapters)
            {
                var chapterPageEl = HierarchySearchManager.FindChapterPage(oneNoteApp, sectionPagesEl.Root, baseChapter.Index, xnm);

                if (chapterPageEl == null)
                    throw new Exception(string.Format("The page for chapter {0} of book {1} does not found", baseChapter.Index, baseBookInfo.Name));

                string chapterPageId = (string)chapterPageEl.Attribute("ID");
                var chapterPageDoc = OneNoteUtils.GetPageContent(oneNoteApp, chapterPageId, out xnm);

                var tableEl = NotebookGenerator.GetBibleTable(chapterPageDoc, xnm);
                int bibleIndex = NotebookGenerator.ExtendBibleTableForParallelTranslation(tableEl, SettingsManager.Instance.PageWidth_Bible, xnm);
                                           
                foreach (var baseVerse in baseChapter.Verses)
                {
                    var baseVersePointer = new SimpleVersePointer(baseBibleBook.Index, baseChapter.Index, baseVerse.Index);

                    var parallelVerse = GetParallelVerse(baseVersePointer, parallelBibleBook, partVersesAlphabet, bookVersePointersComparisonTable,
                                                                                                lastProcessedChapter, lastProcessedVerse, result);

                    NotebookGenerator.AddParallelVerseRowToBibleTable(tableEl, parallelVerse, bibleIndex, baseVersePointer, locale, xnm);

                    lastProcessedChapter = parallelVerse.Chapter;
                    lastProcessedVerse = parallelVerse.TopVerse ?? parallelVerse.Verse;
                }             

                oneNoteApp.UpdatePageContent(chapterPageDoc.ToString(), DateTime.MinValue, Constants.CurrentOneNoteSchema);
            }

            return result;
        }        

        private static SimpleVerse GetParallelVerse(SimpleVersePointer baseVersePointer, BibleBookContent parallelBibleBook, string partVersesAlphabet,
            SimpleVersePointersComparisonTable bookVersePointersComparisonTable, int lastProcessedChapter, int lastProcessedVerse, 
            List<BaseVersePointerException> errors)
        {
            ComparisonVersesInfo parallelVersePointers = null;
            SimpleVersePointer firstParallelVerse = null;

            try
            {
                parallelVersePointers = bookVersePointersComparisonTable.ContainsKey(baseVersePointer)
                                                        ? bookVersePointersComparisonTable[baseVersePointer]
                                                        : new ComparisonVersesInfo { baseVersePointer };                

                if (parallelVersePointers.Count == 0)
                    throw new GetParallelVerseException("parallelVersePointers.Count == 0", baseVersePointer, BaseVersePointerException.Severity.Error);

                firstParallelVerse = parallelVersePointers.First();

                CheckVerseForWarnings(baseVersePointer, parallelBibleBook, firstParallelVerse, lastProcessedChapter, lastProcessedVerse, errors);                

                return GetParallelVerses(baseVersePointer, parallelVersePointers, parallelBibleBook, partVersesAlphabet);
            }
            catch (BaseVersePointerException ex)
            {
                errors.Add(ex);
                return new SimpleVerse(firstParallelVerse != null ? firstParallelVerse : baseVersePointer, string.Empty);
            }
        }

        private static void CheckVerseForWarnings(SimpleVersePointer baseVersePointer, BibleBookContent parallelBibleBook,
            SimpleVersePointer firstParallelVerse, int lastProcessedChapter, int lastProcessedVerse, List<BaseVersePointerException> errors)
        {
            try
            {
                if (lastProcessedChapter > 0 && firstParallelVerse.Chapter > lastProcessedChapter)
                {
                    if (parallelBibleBook.Chapters[lastProcessedChapter - 1].Verses.Count > lastProcessedVerse)
                    {
                        throw new GetParallelVerseException("Miss verse (x01)", baseVersePointer, BaseVersePointerException.Severity.Warning);
                        //parallelVersePointer = new SimpleVersePointer(baseVersePointer.BookIndex, lastProcessedChapter, lastProcessedVerse + 1);
                    }
                    else if (firstParallelVerse.Verse > 1)  // начали главу не с начала
                    {
                        throw new GetParallelVerseException("Miss verse (x02)", baseVersePointer, BaseVersePointerException.Severity.Warning);
                    }
                }
                else if (lastProcessedVerse > 0 && firstParallelVerse.Verse > lastProcessedVerse + 1)
                {
                    throw new GetParallelVerseException("Miss verse (x03)", baseVersePointer, BaseVersePointerException.Severity.Warning);
                    //parallelVersePointer = new SimpleVersePointer(baseVersePointer.BookIndex, lastProcessedChapter, lastProcessedChapter + 1);
                }
                else if (lastProcessedChapter == firstParallelVerse.Chapter && lastProcessedVerse == firstParallelVerse.Verse && !firstParallelVerse.PartIndex.HasValue)
                {
                    throw new GetParallelVerseException("Double verse (x04)", baseVersePointer, BaseVersePointerException.Severity.Warning);
                }
            }
            catch (BaseVersePointerException ex)
            {
                errors.Add(ex);
            }
        }

        private static SimpleVerse GetParallelVerses(SimpleVersePointer baseVersePointer,
            ComparisonVersesInfo parallelVersePointers, BibleBookContent parallelBibleBook, string partVersesAlphabet)
        {
            string verseContent = string.Empty;

            var firstParallelVerse = parallelVersePointers.First();
            int? topVerse = null;

            if (parallelVersePointers[0].Chapter != baseVersePointer.Chapter)
                    verseContent = string.Format("{0}:{1} ", firstParallelVerse.Chapter, firstParallelVerse.Verse);
                else
                    verseContent = string.Format("{0}", firstParallelVerse.Verse);

            if (parallelVersePointers.Count > 1)
            {
                verseContent += string.Format("-{0} {1}", parallelVersePointers.Last().Verse, 
                    parallelBibleBook.GetVersesContent(parallelVersePointers));
                topVerse = parallelVersePointers.Last().Verse;
            }
            else
            {
                string text = parallelBibleBook.GetVerseContent(firstParallelVerse);

                if (string.IsNullOrEmpty(text))  // значит нет такого стиха, либо такой по счёту части стиха
                {
                    if (parallelVersePointers.Strict)
                    {
                        throw new GetParallelVerseException(
                            string.Format("Can not found verseContent (versePart = {0})", firstParallelVerse.PartIndex), baseVersePointer, BaseVersePointerException.Severity.Warning);
                    }
                    else
                        verseContent = string.Empty;
                }
                else
                {
                    if (firstParallelVerse.PartIndex.HasValue)
                    {
                        if (string.IsNullOrEmpty(partVersesAlphabet) || partVersesAlphabet.Length <= firstParallelVerse.PartIndex.Value)
                            partVersesAlphabet = Consts.Constants.DefaultPartVersesAlphabet;

                        verseContent += string.Format("({0})", partVersesAlphabet[firstParallelVerse.PartIndex.Value]);
                    }
                    verseContent += string.Format(" {0}", text);
                }
            }

            return new SimpleVerse(firstParallelVerse, verseContent) { TopVerse = topVerse };            
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
