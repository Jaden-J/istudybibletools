﻿using System;
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
using BibleCommon.Contracts;
using BibleCommon.Scheme;

namespace BibleCommon.Services
{
    public class BibleParallelTranslationConnectionResult
    {
        public List<Exception> Errors { get; set; }

        public BibleParallelTranslationConnectionResult()
        {
            this.Errors = new List<Exception>();
        }
    }

    public class BibleIteratorArgs
    {
        public XDocument ChapterDocument { get; set; }
        public XElement TableElement { get; set; }
        public int BibleIndex { get; set; }
        public int StrongStyleIndex { get; set; }
        public string StrongPrefix { get; set; }
    }

    public class BibleParallelTranslationManager : IDisposable
    {
        public static readonly Version SupportedModuleMinVersion = new Version(2, 0);

        private Application _oneNoteApp;

        public string BibleNotebookId { get; set; }
        public string BaseModuleShortName { get; set; }
        public string ParallelModuleShortName { get; set; }

        public ModuleInfo BaseModuleInfo { get; set; }
        public ModuleInfo ParallelModuleInfo { get; set; }

        public XMLBIBLE BaseBibleInfo { get; set; }
        public XMLBIBLE ParallelBibleInfo { get; set; }

        public ICustomLogger Logger { get; set; }

        public List<Exception> Errors { get; set; }
        public bool ForCheckOnly { get; set; }

        public BibleParallelTranslationManager(Application oneNoteApp, string baseModuleShortName, string parallelModuleShortName, string bibleNotebookId)
        {            
            this.BibleNotebookId = bibleNotebookId;
            this.BaseModuleShortName = baseModuleShortName;
            this.ParallelModuleShortName = parallelModuleShortName;

            this.BaseModuleInfo = ModulesManager.GetModuleInfo(this.BaseModuleShortName);
            this.ParallelModuleInfo = ModulesManager.GetModuleInfo(this.ParallelModuleShortName);

            this.BaseBibleInfo = ModulesManager.GetModuleBibleInfo(this.BaseModuleShortName);
            this.ParallelBibleInfo = ModulesManager.GetModuleBibleInfo(this.ParallelModuleShortName);

            Errors = new List<Exception>();

            CheckModules();

            _oneNoteApp = oneNoteApp;
        }

        public static bool IsModuleSupported(ModuleInfo moduleInfo)
        {
            return moduleInfo.Version >= SupportedModuleMinVersion;
        }

        private void CheckModules()
        {
            ModulesManager.CheckModule(BaseModuleInfo);
            ModulesManager.CheckModule(ParallelModuleInfo);

            if (!IsModuleSupported(BaseModuleInfo))
                throw new NotSupportedException(string.Format("Version of base module is {0}. Only {1} and greater is supported.", BaseModuleInfo.Version, SupportedModuleMinVersion));

            if (!IsModuleSupported(ParallelModuleInfo))
                throw new NotSupportedException(string.Format("Version of parallel module is {0}. Only {1} and greater is supported.", ParallelModuleInfo.Version, SupportedModuleMinVersion));
        }

        public void Dispose()
        {
            _oneNoteApp = null;
        }

        public void RemoveParallelTranslation(string moduleName)
        {
            var moduleInfo = ModulesManager.GetModuleInfo(moduleName);

            XmlNamespaceManager xnm = OneNoteUtils.GetOneNoteXNM();

            IterateBaseBible(
                (chapterPageDoc, chapterPointer) =>
                {
                    RemoveChapterParallelTranslation(chapterPageDoc, moduleInfo, xnm);

                    return null;
                }, true, false, null);

        }

        internal static void RemoveChapterParallelTranslation(XDocument chapterPageDoc, ModuleInfo lastModuleInfo, XmlNamespaceManager xnm)
        {
            var tableEl = NotebookGenerator.GetPageTable(chapterPageDoc, xnm);

            var cellIndex = 0;
            var cellFound = false;
            foreach (var cell in tableEl.XPathSelectElements("one:Row[1]/one:Cell/one:OEChildren/one:OE/one:T", xnm))
            {
                if (StringUtils.GetText(cell.Value) == lastModuleInfo.Name)
                {
                    cellFound = true;
                    break;
                }
                cellIndex++;
            }

            if (cellFound)
            {
                tableEl.XPathSelectElements(string.Format("one:Row/one:Cell[{0}]", cellIndex + 1), xnm).Remove();
                tableEl.XPathSelectElements(string.Format("one:Columns/one:Column[{0}]", cellIndex + 1), xnm).Remove();
            }

            int index = 0;
            foreach (var column in tableEl.XPathSelectElements("one:Columns/one:Column", xnm))
            {
                column.SetAttributeValue("index", index++);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="chapterAction"></param>
        /// <param name="chapterUndoAction">если например ни одного стиха в главе не оказалось и ничего не сделали - то чтобы удалить заготовку колонки</param>
        /// <param name="needToUpdateChapter"></param>
        /// <param name="iterateVerses"></param>
        /// <param name="verseAction"></param>
        /// <returns></returns>
        public BibleParallelTranslationConnectionResult IterateBaseBible(
            Func<XDocument, SimpleVersePointer, BibleIteratorArgs> chapterAction, bool needToUpdateChapter, 
            bool iterateVerses, Action<SimpleVersePointer, SimpleVerse, BibleIteratorArgs> verseAction)
        {
            Errors.Clear();

            var bibleVersePointersComparisonTable = BibleParallelTranslationConnectorManager.GetParallelBibleInfo(
                                                          BaseModuleInfo.ShortName, ParallelModuleInfo.ShortName,
                                                          BaseModuleInfo.BibleTranslationDifferences,
                                                          ParallelModuleInfo.BibleTranslationDifferences);

            var result = new BibleParallelTranslationConnectionResult();

            foreach (var baseBookContent in BaseBibleInfo.Books)
            {
                var baseBookInfo = BaseModuleInfo.BibleStructure.BibleBooks.FirstOrDefault(b => b.Index == baseBookContent.Index);
                if (baseBookInfo == null)
                    throw new InvalidModuleException(string.Format("Book with index {0} is not found in module manifest", baseBookContent.Index));                

                var parallelBookContent = ParallelBibleInfo.Books.FirstOrDefault(b => b.Index == baseBookContent.Index);
                if (parallelBookContent != null)
                {
                    XElement sectionEl = ForCheckOnly ? null : HierarchySearchManager.FindBibleBookSection(_oneNoteApp, BibleNotebookId, baseBookInfo.SectionName);
                    if (sectionEl == null && !ForCheckOnly)
                        throw new Exception(string.Format("Section with name {0} is not found", baseBookInfo.SectionName));

                    SimpleVersePointersComparisonTable bookVersePointersComparisonTable = bibleVersePointersComparisonTable.ContainsKey(baseBookContent.Index)
                        ? bibleVersePointersComparisonTable[baseBookContent.Index] : new SimpleVersePointersComparisonTable();

                    try
                    {
                        ProcessBibleBook(sectionEl, baseBookInfo, baseBookContent, parallelBookContent, bookVersePointersComparisonTable,
                            chapterAction, needToUpdateChapter, iterateVerses, verseAction);
                    }
                    catch (BaseVersePointerException ex) 
                    {
                        Errors.Add(ex);
                    }
                }
            }

            result.Errors = Errors;

            return result;
        }      

        private void ProcessBibleBook(XElement bibleBookSectionEl, BibleBookInfo baseBookInfo,
            BIBLEBOOK baseBookContent, BIBLEBOOK parallelBookContent, 
            SimpleVersePointersComparisonTable bookVersePointersComparisonTable,
            Func<XDocument, SimpleVersePointer, BibleIteratorArgs> chapterAction, bool needToUpdateChapter,
            bool iterateVerses, Action<SimpleVersePointer, SimpleVerse, BibleIteratorArgs> verseAction)
        {
            XmlNamespaceManager xnm = OneNoteUtils.GetOneNoteXNM();
            string sectionId = ForCheckOnly ? null : (string)bibleBookSectionEl.Attribute("ID");            

            var sectionPagesEl = ForCheckOnly ? null : OneNoteUtils.GetHierarchyElement(_oneNoteApp, sectionId, HierarchyScope.hsPages, out xnm);

            int lastProcessedChapter = 0;
            int lastProcessedVerse = 0;

            foreach (var baseChapter in baseBookContent.Chapters)
            {
                if (Logger != null)                
                    Logger.LogMessage("{0} '{1} {2}'", BibleCommon.Resources.Constants.ProcessChapter, baseBookInfo.Name, baseChapter.Index);                

                XDocument chapterPageDoc = null;
                BibleIteratorArgs bibleIteratorArgs = null;

                if (chapterAction != null)
                {
                    var chapterPageEl = ForCheckOnly ? null : HierarchySearchManager.FindChapterPage(_oneNoteApp, sectionPagesEl.Root, baseChapter.Index, xnm);

                    if (chapterPageEl == null && !ForCheckOnly)
                        throw new BaseChapterSectionNotFoundException(baseChapter.Index, baseBookInfo.Index);

                    string chapterPageId = ForCheckOnly ? null : (string)chapterPageEl.Attribute("ID");
                    chapterPageDoc = ForCheckOnly ? null : OneNoteUtils.GetPageContent(_oneNoteApp, chapterPageId, out xnm);

                    bibleIteratorArgs = chapterAction(chapterPageDoc, new SimpleVersePointer(baseBookInfo.Index, baseChapter.Index));
                }


                bool? chapterWasModified = null;
                if (iterateVerses)
                {
                    chapterWasModified = false;
                    foreach (var baseVerse in baseChapter.Verses)
                    {                        
                        var baseVersePointer = new SimpleVersePointer(baseBookContent.Index, baseChapter.Index, baseVerse.VerseNumber);                        

                        //var originalVersePointer = bookVersePointersComparisonTable.GetOriginalKey(baseVersePointer);
                        //if (originalVersePointer != null && originalVersePointer.IsEmpty)
                        //    continue;                            

                        var parallelVerse = GetParallelVerse(baseVersePointer, parallelBookContent, bookVersePointersComparisonTable, 
                            bibleIteratorArgs != null ? bibleIteratorArgs.StrongPrefix : null,
                            lastProcessedChapter, lastProcessedVerse);

                        if (verseAction != null)
                        {
                            try
                            {
                                verseAction(baseVersePointer, parallelVerse, bibleIteratorArgs);
                            }
                            catch (BaseVersePointerException ex)
                            {
                                Errors.Add(ex);
                            }
                        }

                        if (!parallelVerse.IsEmpty)
                        {
                            lastProcessedChapter = parallelVerse.Chapter;
                            lastProcessedVerse = parallelVerse.TopVerse ?? parallelVerse.Verse;
                            chapterWasModified = true;
                        }
                    }
                }

                if (needToUpdateChapter && chapterAction != null && chapterWasModified.GetValueOrDefault(true) == true && !ForCheckOnly)
                {
                    SupplementalBibleManager.UpdatePageXmlForStrongBible(chapterPageDoc);

                    _oneNoteApp.UpdatePageContent(chapterPageDoc.ToString(), DateTime.MinValue, Constants.CurrentOneNoteSchema);
                }
            }            
        }

        private SimpleVerse GetParallelVerse(SimpleVersePointer baseVersePointer, BIBLEBOOK parallelBookContent, 
            SimpleVersePointersComparisonTable bookVersePointersComparisonTable, string strongPrefix, int lastProcessedChapter, int lastProcessedVerse)
        {
            ComparisonVersesInfo parallelVersePointers = new ComparisonVersesInfo();;
            SimpleVersePointer firstParallelVerse = null;

            try
            {
                baseVersePointer.GetAllVerses().ForEach(
                verse =>
                {
                    var comparisonTable = bookVersePointersComparisonTable.ContainsKey(verse)
                                                        ? bookVersePointersComparisonTable[verse]
                                                        : new ComparisonVersesInfo { verse };
                    comparisonTable.ForEach(pVerse => parallelVersePointers.Add(pVerse));
                });
                    

                if (parallelVersePointers.Count == 0)
                    throw new GetParallelVerseException("parallelVersePointers.Count == 0", baseVersePointer, BaseVersePointerException.Severity.Error);
                
                var parallelVerse = GetParallelVerses(baseVersePointer, parallelVersePointers, parallelBookContent, strongPrefix);
                
                if (!parallelVerse.IsEmpty)
                    CheckVerseForWarnings(baseVersePointer, parallelBookContent, parallelVersePointers.First(), lastProcessedChapter, lastProcessedVerse);  

                return parallelVerse;
            }
            catch (BaseVersePointerException ex)
            {
                if (ex.IsChapterException)
                    throw;

                Errors.Add(ex);
                return new SimpleVerse(firstParallelVerse != null ? firstParallelVerse : baseVersePointer, string.Empty);
            }
        }

        private void CheckVerseForWarnings(SimpleVersePointer baseVersePointer, BIBLEBOOK parallelBookContent,
            SimpleVersePointer firstParallelVerse, int lastProcessedChapter, int lastProcessedVerse)
        {
            try
            {
                if (firstParallelVerse.SkipCheck)
                    return;
                if (!firstParallelVerse.IsEmpty)
                {
                    if (lastProcessedChapter > 0 && firstParallelVerse.Chapter > lastProcessedChapter)
                    {                        
                        if (parallelBookContent.Chapters.Count > lastProcessedChapter - 1)
                        {
                            var previousProcessedChapterlastVerse = parallelBookContent.Chapters[lastProcessedChapter - 1].Verses.Last();
                            if ((previousProcessedChapterlastVerse.TopIndex ?? previousProcessedChapterlastVerse.Index) > lastProcessedVerse)
                                throw new GetParallelVerseException("Miss verse (x01)", baseVersePointer, BaseVersePointerException.Severity.Warning);
                        }
                        else if (firstParallelVerse.Verse > 1)  // начали главу не с начала                    
                            throw new GetParallelVerseException("Miss verse (x02)", baseVersePointer, BaseVersePointerException.Severity.Warning);
                    }
                    else
                    {
                        if (lastProcessedVerse > 0 && firstParallelVerse.Verse > lastProcessedVerse + 1)
                            throw new GetParallelVerseException("Miss verse (x03)", baseVersePointer, BaseVersePointerException.Severity.Warning);
                        else if (lastProcessedChapter == firstParallelVerse.Chapter && lastProcessedVerse == firstParallelVerse.Verse && !firstParallelVerse.PartIndex.HasValue)
                            throw new GetParallelVerseException("Double verse (x04)", baseVersePointer, BaseVersePointerException.Severity.Warning);
                        else if (lastProcessedChapter == firstParallelVerse.Chapter && firstParallelVerse.Verse < lastProcessedVerse)
                            throw new GetParallelVerseException("Reverse verse (x05)", baseVersePointer, BaseVersePointerException.Severity.Warning);                        
                    }
                }
            }
            catch (BaseVersePointerException ex)
            {
                Errors.Add(ex);
            }
        }

        private SimpleVerse GetParallelVerses(SimpleVersePointer baseVersePointer,
            ComparisonVersesInfo parallelVersePointers, BIBLEBOOK parallelBookContent, string strongPrefix)
        {
            string verseContent = string.Empty;
            string verseNumberContent = string.Empty;

            var firstParallelVerse = parallelVersePointers.First();
            int? topLastVerse = null;
            bool isEmpty = false;

            if (!firstParallelVerse.IsEmpty)
            {
                List<SimpleVersePointer> notFoundVerses;
                verseContent = parallelBookContent.GetVersesContent(parallelVersePointers, strongPrefix, out topLastVerse, out isEmpty, out notFoundVerses);                
                if (!isEmpty)
                {
                    verseNumberContent = GetVersesNumberString(baseVersePointer, parallelVersePointers, topLastVerse);                  

                    if (verseContent == null)
                    {
                        throw new GetParallelVerseException(                                // значит нет такого стиха, либо такой по счёту части стиха      
                            string.Format("Can not find verseContent{0}",
                                            firstParallelVerse.PartIndex.HasValue
                                                ? string.Format(" (versePart = {0})", firstParallelVerse.PartIndex + 1)
                                                : string.Empty),
                                            baseVersePointer, BaseVersePointerException.Severity.Warning);
                    }
                    else  
                    {
                        foreach (var notFoundVerse in notFoundVerses)
                        {
                            Errors.Add(new GetParallelVerseException(                        // значит один из нескольких не удалось найти
                                string.Format("Can not find verseContent{0}",
                                                notFoundVerse.PartIndex.HasValue
                                                    ? string.Format(" (versePart = {0})", notFoundVerse.PartIndex + 1)
                                                    : string.Empty),
                                                baseVersePointer, BaseVersePointerException.Severity.Warning));
                        }
                    }
                }                
            }

            return new SimpleVerse(firstParallelVerse, verseNumberContent, verseContent) 
            {
                VerseNumber = new VerseNumber(firstParallelVerse.Verse, topLastVerse),
                IsEmpty = firstParallelVerse.IsEmpty || isEmpty
            };
        }

        private string GetVersesNumberString(SimpleVersePointer baseVersePointer, ComparisonVersesInfo parallelVersePointers, int? topVerse)
        {
            string result = string.Empty;
            var firstParallelVerse = parallelVersePointers.First();

            if (!firstParallelVerse.IsEmpty)
            {
                result = GetVerseNumberString(firstParallelVerse, null);

                if (parallelVersePointers[0].Chapter != baseVersePointer.Chapter)
                    result = string.Format("{0}:{1}", firstParallelVerse.Chapter, result);

                if (parallelVersePointers.Count > 1 || topVerse.HasValue)
                {
                    var lastVerse = parallelVersePointers.Last();

                    result += string.Format("-{0}", GetVerseNumberString(lastVerse, topVerse));
                }
            }

            return result;
        }

        private string GetVerseNumberString(SimpleVersePointer versePointer, int? topVerse)
        {
            var result = topVerse.HasValue ? topVerse.ToString() : versePointer.VerseNumber.ToString();
            if (versePointer.PartIndex.HasValue)
            {
                var partVersesAlphabet = ParallelModuleInfo.BibleTranslationDifferences.PartVersesAlphabet;
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
        
        public static void MergeModuleWithMainBible(string parallelModuleName)
        {
            if (SettingsManager.Instance.ModuleName != parallelModuleName)
            {
                var baseModuleInfo = ModulesManager.GetModuleInfo(SettingsManager.Instance.ModuleName);
                var parallelModuleInfo = ModulesManager.GetModuleInfo(parallelModuleName);                

                // merge book abbriviations
                foreach (var baseBook in baseModuleInfo.BibleStructure.BibleBooks)
                {
                    var parallelBook = parallelModuleInfo.BibleStructure.BibleBooks.FirstOrDefault(b => b.Index == baseBook.Index);
                    if (parallelBook != null)
                    {
                        foreach (var parallelBookAbbreviation in parallelBook.Abbreviations)
                        {
                            if (!baseBook.Abbreviations.Exists(abbr => abbr.Value == parallelBookAbbreviation.Value))
                            {
                                baseBook.Abbreviations.Add(new Abbreviation(parallelBookAbbreviation.Value)
                                {
                                    ModuleName = parallelModuleName,
                                    IsFullBookName = parallelBookAbbreviation.IsFullBookName
                                });
                            }
                        }
                    }
                }

                //merge alphabets
                foreach (var c in parallelModuleInfo.BibleStructure.Alphabet)
                {
                    if (!baseModuleInfo.BibleStructure.Alphabet.Contains(c))
                        baseModuleInfo.BibleStructure.Alphabet += c;
                }

                ModulesManager.UpdateModuleManifest(baseModuleInfo);
            }
        }
        
        public static void RemoveBookAbbreviationsFromMainBible(string parallelModuleName)
        {
            if (SettingsManager.Instance.ModuleName != parallelModuleName)
            {
                var baseModuleInfo = ModulesManager.GetModuleInfo(SettingsManager.Instance.ModuleName);

                foreach (var baseBook in baseModuleInfo.BibleStructure.BibleBooks)
                {
                    baseBook.Abbreviations.RemoveAll(abbr => abbr.ModuleName == parallelModuleName);
                }

                ModulesManager.UpdateModuleManifest(baseModuleInfo);
            }
        }
    }
}
