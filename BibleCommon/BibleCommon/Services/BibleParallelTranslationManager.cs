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
using BibleCommon.Contracts;

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

    public class BibleIteratorArgs
    {
        public XDocument ChapterDocument { get; set; }
        public XElement TableElement { get; set; }
        public int BibleIndex { get; set; }        
    }

    public class BibleParallelTranslationManager : IDisposable
    {
        public const string SupportedModuleMinVersion = "2.0";

        private Application _oneNoteApp;

        public string BibleNotebookId { get; set; }
        public string BaseModuleShortName { get; set; }
        public string ParallelModuleShortName { get; set; }

        public ModuleInfo BaseModuleInfo { get; set; }
        public ModuleInfo ParallelModuleInfo { get; set; }

        public ModuleBibleInfo BaseBibleInfo { get; set; }
        public ModuleBibleInfo ParallelBibleInfo { get; set; }

        public ICustomLogger Logger { get; set; }

        public List<BaseVersePointerException> Errors { get; set; }

        public BibleParallelTranslationManager(Application oneNoteApp, string baseModuleShortName, string parallelModuleShortName, string bibleNotebookId)
        {            
            this.BibleNotebookId = bibleNotebookId;
            this.BaseModuleShortName = baseModuleShortName;
            this.ParallelModuleShortName = parallelModuleShortName;

            this.BaseModuleInfo = ModulesManager.GetModuleInfo(this.BaseModuleShortName);
            this.ParallelModuleInfo = ModulesManager.GetModuleInfo(this.ParallelModuleShortName);

            this.BaseBibleInfo = ModulesManager.GetModuleBibleInfo(this.BaseModuleShortName);
            this.ParallelBibleInfo = ModulesManager.GetModuleBibleInfo(this.ParallelModuleShortName);

            Errors = new List<BaseVersePointerException>();

            CheckModules();

            _oneNoteApp = oneNoteApp;
        }

        public static bool IsModuleSupported(ModuleInfo moduleInfo)
        {
            return moduleInfo.Version.CompareTo(SupportedModuleMinVersion) >= 0;
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

        public void RemoveLastParallelTranslation()
        {
            if (SettingsManager.Instance.SupplementalBibleModules.Count > 1)
            {
                string lastModuleName = SettingsManager.Instance.SupplementalBibleModules.Last();
                var moduleInfo = ModulesManager.GetModuleInfo(lastModuleName);

                XmlNamespaceManager xnm = OneNoteUtils.GetOneNoteXNM();

                IterateBaseBible(chapterPageDoc =>
                {
                    var tableEl = NotebookGenerator.GetBibleTable(chapterPageDoc, xnm);

                    var cellIndex = 0;
                    var cellFound = false;
                    foreach (var cell in tableEl.XPathSelectElements("one:Row[1]/one:Cell/one:OEChildren/one:OE/one:T", xnm))
                    {
                        if (StringUtils.GetText(cell.Value) == moduleInfo.Name)
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

                    return null;
                }, true, false, null);
            }
        }
        
        public BibleParallelTranslationConnectionResult AddParallelTranslation()
        {
            XmlNamespaceManager xnm = OneNoteUtils.GetOneNoteXNM();            
            
            return IterateBaseBible(chapterPageDoc =>
                {
                    var tableEl = NotebookGenerator.GetBibleTable(chapterPageDoc, xnm);
                    var bibleIndex = NotebookGenerator.AddColumnToTable(tableEl, SettingsManager.Instance.PageWidth_Bible, xnm);
                    NotebookGenerator.AddParallelBibleTitle(tableEl, ParallelModuleInfo.Name, bibleIndex, ParallelBibleInfo.Content.Locale, xnm);

                    return new BibleIteratorArgs() { BibleIndex = bibleIndex, TableElement = tableEl };
                }, true, true,
                (baseVersePointer, parallelVerse, bibleIteratorArgs) =>
                {
                    NotebookGenerator.AddParallelVerseRowToBibleTable(bibleIteratorArgs.TableElement, parallelVerse, 
                        bibleIteratorArgs.BibleIndex, baseVersePointer, ParallelBibleInfo.Content.Locale, xnm);
                });
        }

        public BibleParallelTranslationConnectionResult IterateBaseBible(Func<XDocument, BibleIteratorArgs> chapterAction, bool needToUpdateChapter, 
            bool iterateVerses, Action<SimpleVersePointer, SimpleVerse, BibleIteratorArgs> verseAction)
        {
            Errors.Clear();

            var bibleVersePointersComparisonTable = BibleParallelTranslationConnectorManager.GetParallelBibleInfo(
                                                          BaseModuleInfo.ShortName, ParallelModuleInfo.ShortName,
                                                          BaseModuleInfo.BibleTranslationDifferences,
                                                          ParallelModuleInfo.BibleTranslationDifferences);

            var result = new BibleParallelTranslationConnectionResult();

            foreach (var baseBookContent in BaseBibleInfo.Content.Books)
            {
                var baseBookInfo = BaseModuleInfo.BibleStructure.BibleBooks.FirstOrDefault(b => b.Index == baseBookContent.Index);
                if (baseBookInfo == null)
                    throw new InvalidModuleException(string.Format("Book with index {0} is not found in module manifest", baseBookContent.Index));                

                var parallelBookContent = ParallelBibleInfo.Content.Books.FirstOrDefault(b => b.Index == baseBookContent.Index);
                if (parallelBookContent != null)
                {
                    XElement sectionEl = HierarchySearchManager.FindBibleBookSection(_oneNoteApp, BibleNotebookId, baseBookInfo.SectionName);
                    if (sectionEl == null)
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
            BibleBookContent baseBookContent, BibleBookContent parallelBookContent, 
            SimpleVersePointersComparisonTable bookVersePointersComparisonTable,
            Func<XDocument, BibleIteratorArgs> chapterAction, bool needToUpdateChapter,
            bool iterateVerses, Action<SimpleVersePointer, SimpleVerse, BibleIteratorArgs> verseAction)
        {
            XmlNamespaceManager xnm;
            string sectionId = (string)bibleBookSectionEl.Attribute("ID");
            string sectionName = (string)bibleBookSectionEl.Attribute("name");

            var sectionPagesEl = OneNoteUtils.GetHierarchyElement(_oneNoteApp, sectionId, HierarchyScope.hsPages, out xnm);

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
                    var chapterPageEl = HierarchySearchManager.FindChapterPage(_oneNoteApp, sectionPagesEl.Root, baseChapter.Index, xnm);

                    if (chapterPageEl == null)
                        throw new BaseChapterSectionNotFoundException(baseChapter.Index, baseBookInfo.Index);

                    string chapterPageId = (string)chapterPageEl.Attribute("ID");
                    chapterPageDoc = OneNoteUtils.GetPageContent(_oneNoteApp, chapterPageId, out xnm);

                    bibleIteratorArgs = chapterAction(chapterPageDoc);
                }

                if (iterateVerses)
                {
                    foreach (var baseVerse in baseChapter.Verses)
                    {                        
                        var baseVersePointer = new SimpleVersePointer(baseBookContent.Index, baseChapter.Index, baseVerse.Index);

                        if (bookVersePointersComparisonTable.ContainsKey(baseVersePointer) && bookVersePointersComparisonTable[baseVersePointer]  здесь бы понять, что он IsEmpty.
                            

                        var parallelVerse = GetParallelVerse(baseVersePointer, parallelBookContent, bookVersePointersComparisonTable, lastProcessedChapter, lastProcessedVerse);

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
                        }
                    }
                }

                if (needToUpdateChapter && chapterAction != null)
                    _oneNoteApp.UpdatePageContent(chapterPageDoc.ToString(), DateTime.MinValue, Constants.CurrentOneNoteSchema);
            }            
        }       

        private  SimpleVerse GetParallelVerse(SimpleVersePointer baseVersePointer, BibleBookContent parallelBookContent, 
            SimpleVersePointersComparisonTable bookVersePointersComparisonTable, int lastProcessedChapter, int lastProcessedVerse)
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

                CheckVerseForWarnings(baseVersePointer, parallelBookContent, firstParallelVerse, lastProcessedChapter, lastProcessedVerse);

                return GetParallelVerses(baseVersePointer, parallelVersePointers, parallelBookContent);
            }
            catch (BaseVersePointerException ex)
            {
                if (ex.IsChapterException)
                    throw;

                Errors.Add(ex);
                return new SimpleVerse(firstParallelVerse != null ? firstParallelVerse : baseVersePointer, string.Empty);
            }
        }

        private void CheckVerseForWarnings(SimpleVersePointer baseVersePointer, BibleBookContent parallelBookContent,
            SimpleVersePointer firstParallelVerse, int lastProcessedChapter, int lastProcessedVerse)
        {
            try
            {
                if (lastProcessedChapter > 0 && firstParallelVerse.Chapter > lastProcessedChapter)
                {
                    if ((parallelBookContent.Chapters.Count > lastProcessedChapter - 1) && (parallelBookContent.Chapters[lastProcessedChapter - 1].Verses.Count > lastProcessedVerse))
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
                Errors.Add(ex);
            }
        }

        private SimpleVerse GetParallelVerses(SimpleVersePointer baseVersePointer,
            ComparisonVersesInfo parallelVersePointers, BibleBookContent parallelBookContent)
        {
            string verseContent = string.Empty;
            string verseNumberContent = string.Empty;

            var firstParallelVerse = parallelVersePointers.First();
            int? topVerse = null;

            if (!firstParallelVerse.IsEmpty)
            {
                verseNumberContent = GetVersesNumberString(baseVersePointer, parallelVersePointers);
                verseContent = parallelBookContent.GetVersesContent(parallelVersePointers);

                if (string.IsNullOrEmpty(verseContent))  // значит нет такого стиха, либо такой по счёту части стиха                                    
                    throw new GetParallelVerseException(
                        string.Format("Can not find verseContent (versePart = {0})", firstParallelVerse.PartIndex + 1), baseVersePointer, BaseVersePointerException.Severity.Warning);

                if (parallelVersePointers.Count > 1)
                    topVerse = parallelVersePointers.Last().Verse;
            }

            return new SimpleVerse(firstParallelVerse, string.Format("{0}{1}{2}",
                                                            verseNumberContent,
                                                            string.IsNullOrEmpty(verseContent) ? string.Empty : " ",
                                                            verseContent)) { TopVerse = topVerse, IsEmpty = firstParallelVerse.IsEmpty };
        }

        private string GetVersesNumberString(SimpleVersePointer baseVersePointer, ComparisonVersesInfo parallelVersePointers)
        {
            string result = string.Empty;
            var firstParallelVerse = parallelVersePointers.First();

            if (!firstParallelVerse.IsEmpty)
            {
                result = GetVerseNumberString(firstParallelVerse);

                if (parallelVersePointers[0].Chapter != baseVersePointer.Chapter)
                    result = string.Format("{0}:{1} ", firstParallelVerse.Chapter, result);

                if (parallelVersePointers.Count > 1)
                {
                    var topVerse = parallelVersePointers.Last();

                    result += string.Format("-{0}", GetVerseNumberString(topVerse));
                }
            }

            return result;
        }

        private string GetVerseNumberString(SimpleVersePointer versePointer)
        {
            string partVersesAlphabet = ParallelModuleInfo.BibleTranslationDifferences.PartVersesAlphabet;
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

        /// <summary>
        /// With base Bible
        /// </summary>
        /// <param name="parallelModuleName"></param>
        public static void AggregateBookAbbreviations(string parallelModuleName)
        {
            if (SettingsManager.Instance.ModuleName != parallelModuleName)
            {
                var baseModuleInfo = ModulesManager.GetModuleInfo(SettingsManager.Instance.ModuleName);
                var parallelModuleInfo = ModulesManager.GetModuleInfo(parallelModuleName);

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

                ModulesManager.UpdateModuleManifest(baseModuleInfo);
            }
        }

        /// <summary>
        /// From base Bible
        /// </summary>
        /// <param name="parallelModuleName"></param>
        public static void RemoveBookAbbreviations(string parallelModuleName)
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
