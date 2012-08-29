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
using System.Xml.XPath;

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

            var baseBibleInfo = ModulesManager.GetModuleBibleInfo(baseModuleInfo.ShortName);
            var parallelBibleInfo = ModulesManager.GetModuleBibleInfo(moduleShortName);

            var bibleVersePointersComparisonTable = BibleParallelTranslationConnectorManager.GetBibleParallelTranslationInfo(
                                                            baseModuleInfo.ShortName, parallelModuleInfo.ShortName,
                                                            baseModuleInfo.BibleTranslationDifferences,
                                                            parallelModuleInfo.BibleTranslationDifferences);                        

            return GenerateParallelBibleTables(oneNoteApp, SettingsManager.Instance.NotebookId_SupplementalBible,
                baseModuleInfo, baseBibleInfo, parallelModuleInfo, parallelBibleInfo, bibleVersePointersComparisonTable);            
        }

        private static BibleParallelTranslationConnectionResult GenerateParallelBibleTables(Application oneNoteApp, string bibleNotebookId, 
            ModuleInfo baseModuleInfo, ModuleBibleInfo baseBibleInfo, 
            ModuleInfo parallelModuleInfo, ModuleBibleInfo parallelBibleInfo,
            ParallelBibleInfo bibleVersePointersComparisonTable)
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
                        parallelBibleBook, parallelModuleInfo.Name, parallelModuleInfo.BibleTranslationDifferences.PartVersesAlphabet, 
                        bookVersePointersComparisonTable, parallelBibleInfo.Content.Locale));                    
                }
            }

            return result;
        }

        private static List<BaseVersePointerException> ProcessBibleBook(Application oneNoteApp, XElement sectionEl, BibleBookInfo baseBibleInfo, 
            BibleBookContent baseBibleBook, BibleBookInfo baseBookInfo, BibleBookContent parallelBibleBook, 
            string parallelTranslationModuleName, string partVersesAlphabet,
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
                int bibleIndex = NotebookGenerator.AddColumnToTable(tableEl, SettingsManager.Instance.PageWidth_Bible, xnm);

                NotebookGenerator.AddParallelBibleTitle(tableEl, parallelTranslationModuleName, bibleIndex, locale, xnm);
                                           
                foreach (var baseVerse in baseChapter.Verses)
                {
                    var baseVersePointer = new SimpleVersePointer(baseBibleBook.Index, baseChapter.Index, baseVerse.Index);

                    var parallelVerse = GetParallelVerse(baseVersePointer, parallelBibleBook, partVersesAlphabet, bookVersePointersComparisonTable,
                                                                                                lastProcessedChapter, lastProcessedVerse, result);

                    NotebookGenerator.AddParallelVerseRowToBibleTable(tableEl, parallelVerse, bibleIndex, baseVersePointer, locale, xnm);

                    if (!parallelVerse.IsEmpty)
                    {
                        lastProcessedChapter = parallelVerse.Chapter;
                        lastProcessedVerse = parallelVerse.TopVerse ?? parallelVerse.Verse;
                    }
                }             

                oneNoteApp.UpdatePageContent(chapterPageDoc.ToString(), DateTime.MinValue, Constants.CurrentOneNoteSchema);
            }

            return result;
        }


        доделать этот метод
        public static void IterateBaseBible(Application oneNoteApp, XElement sectionEl, BibleBookInfo baseBibleInfo,
            BibleBookContent baseBibleBook, BibleBookInfo baseBookInfo, BibleBookContent parallelBibleBook,
            string parallelTranslationModuleName, string partVersesAlphabet,
            SimpleVersePointersComparisonTable bookVersePointersComparisonTable, string locale)
        {
            int lastProcessedChapter = 0;
            int lastProcessedVerse = 0;     

            foreach (var baseChapter in baseBibleBook.Chapters)
            {
                foreach (var baseVerse in baseChapter.Verses)
                {
                    var baseVersePointer = new SimpleVersePointer(baseBibleBook.Index, baseChapter.Index, baseVerse.Index);

                    var parallelVerse = GetParallelVerse(baseVersePointer, parallelBibleBook, partVersesAlphabet, bookVersePointersComparisonTable,
                                                                                                lastProcessedChapter, lastProcessedVerse, result);                    

                    if (!parallelVerse.IsEmpty)
                    {
                        lastProcessedChapter = parallelVerse.Chapter;
                        lastProcessedVerse = parallelVerse.TopVerse ?? parallelVerse.Verse;
                    }
                }       
            }
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
                        throw new GetParallelVerseException("Miss verse (x01)", baseVersePointer, BaseVersePointerException.Severity.Warning);
                    else if (firstParallelVerse.Verse > 1)  // начали главу не с начала                    
                        throw new GetParallelVerseException("Miss verse (x02)", baseVersePointer, BaseVersePointerException.Severity.Warning);
                }
                else
                {
                    if (lastProcessedVerse > 0 && firstParallelVerse.Verse > lastProcessedVerse + 1)                    
                        throw new GetParallelVerseException("Miss verse (x03)", baseVersePointer, BaseVersePointerException.Severity.Warning);                    
                    else if (lastProcessedChapter == firstParallelVerse.Chapter && lastProcessedVerse == firstParallelVerse.Verse && !firstParallelVerse.PartIndex.HasValue)                    
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
            string verseNumberContent = string.Empty;

            var firstParallelVerse = parallelVersePointers.First();
            int? topVerse = null;

            if (!firstParallelVerse.IsEmpty)                            
            {                
                verseNumberContent = GetVersesNumberString(baseVersePointer, parallelVersePointers, partVersesAlphabet);
                verseContent = parallelBibleBook.GetVersesContent(parallelVersePointers);

                if (string.IsNullOrEmpty(verseContent))  // значит нет такого стиха, либо такой по счёту части стиха                                    
                    throw new GetParallelVerseException(
                        string.Format("Can not find verseContent (versePart = {0})", firstParallelVerse.PartIndex + 1), baseVersePointer, BaseVersePointerException.Severity.Warning);

                if (parallelVersePointers.Count > 1)
                    topVerse = parallelVersePointers.Last().Verse;
            }

            return new SimpleVerse(firstParallelVerse, string.Format("{0}{1}{2}", 
                                                            verseNumberContent, 
                                                            string.IsNullOrEmpty(verseContent) ? string.Empty : " ",
                                                            verseContent))
                        { TopVerse = topVerse, IsEmpty = firstParallelVerse.IsEmpty };            
        }

        private static string GetVersesNumberString(SimpleVersePointer baseVersePointer, ComparisonVersesInfo parallelVersePointers, string partVersesAlphabet)
        {
            string result = string.Empty;            
            var firstParallelVerse = parallelVersePointers.First();            

            if (!firstParallelVerse.IsEmpty)
            {
                result = GetVerseNumberString(firstParallelVerse, partVersesAlphabet);

                if (parallelVersePointers[0].Chapter != baseVersePointer.Chapter)
                    result = string.Format("{0}:{1} ", firstParallelVerse.Chapter, result);                

                if (parallelVersePointers.Count > 1)
                {
                    var topVerse = parallelVersePointers.Last();

                    result += string.Format("-{0}", GetVerseNumberString(topVerse, partVersesAlphabet));
                }               
            }

            return result;
        }

        private static string GetVerseNumberString(SimpleVersePointer versePointer, string partVersesAlphabet)
        {
            var result = string.Format("{0}", versePointer.Verse);
            if (versePointer.PartIndex.HasValue)
            {
                if (string.IsNullOrEmpty(partVersesAlphabet) || partVersesAlphabet.Length <= versePointer.PartIndex.Value)
                    partVersesAlphabet = Consts.Constants.DefaultPartVersesAlphabet;

                result += string.Format("({0})", partVersesAlphabet[versePointer.PartIndex.Value]);
            }
            return result;
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
