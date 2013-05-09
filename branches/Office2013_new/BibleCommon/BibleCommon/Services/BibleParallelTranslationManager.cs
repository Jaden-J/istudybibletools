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
using BibleCommon.Scheme;

namespace BibleCommon.Services
{
    public class BibleParallelTranslationConnectionResult
    {
        public List<Exception> Errors { get; set; }
        public List<string> NotFoundBibleVerses { get; set; }

        public BibleParallelTranslationConnectionResult()
        {
            Errors = new List<Exception>();
            NotFoundBibleVerses = new List<string>();
        }
    }

    public class BibleIteratorArgs
    {
        public XDocument ChapterDocument { get; set; }
        public XElement TableElement { get; set; }
        public int BibleIndex { get; set; }
        public int StrongStyleIndex { get; set; }
        public string StrongPrefix { get; set; }
        public bool? NotNeedToUpdateChapter { get; set; }
        public bool? NotNeedToProcessVerses { get; set; }
    }

    public class BibleParallelTranslationManager : IDisposable
    {
        public static readonly Version SupportedModuleMinVersion = new Version(2, 0);

        private Application _oneNoteApp;
        private bool _isOneNote2010;

        public string BibleNotebookId { get; set; }
        public string BaseModuleShortName { get; set; }
        public string ParallelModuleShortName { get; set; }

        public ModuleInfo BaseModuleInfo { get; set; }
        public ModuleInfo ParallelModuleInfo { get; set; }

        public XMLBIBLE BaseBibleInfo { get; set; }
        public XMLBIBLE ParallelBibleInfo { get; set; }

        public ICustomLogger Logger { get; set; }        

        public bool ForCheckOnly { get; set; }

        private BibleParallelTranslationConnectionResult _result;

        public BibleParallelTranslationManager(string baseModuleShortName, string parallelModuleShortName, string bibleNotebookId)
        {            
            this.BibleNotebookId = bibleNotebookId;
            this.BaseModuleShortName = baseModuleShortName;
            this.ParallelModuleShortName = parallelModuleShortName;

            this.BaseModuleInfo = ModulesManager.GetModuleInfo(this.BaseModuleShortName);
            this.ParallelModuleInfo = ModulesManager.GetModuleInfo(this.ParallelModuleShortName);

            this.BaseBibleInfo = ModulesManager.GetModuleBibleInfo(this.BaseModuleShortName);
            this.ParallelBibleInfo = ModulesManager.GetModuleBibleInfo(this.ParallelModuleShortName);            

            CheckModules();

            _oneNoteApp = OneNoteUtils.CreateOneNoteAppSafe();
            _isOneNote2010 = true; // OneNoteUtils.IsOneNote2010Cached(_oneNoteApp);
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
            OneNoteUtils.ReleaseOneNoteApp(ref _oneNoteApp);
        }

        public void RemoveParallelTranslation(string moduleName)
        {
            var moduleInfo = ModulesManager.GetModuleInfo(moduleName);

            XmlNamespaceManager xnm = OneNoteUtils.GetOneNoteXNM();

            IterateBaseBible(
                (chapterPageDoc, chapterPointer) =>
                {
                    return new BibleIteratorArgs() 
                    { 
                        NotNeedToUpdateChapter = !RemoveChapterParallelTranslation(chapterPageDoc, moduleInfo, xnm) 
                    };                    
                }, true, false, null);
        }

        internal static bool RemoveChapterParallelTranslation(XDocument chapterPageDoc, ModuleInfo moduleInfo, XmlNamespaceManager xnm)
        {
            var supplementalModulesMetadata = OneNoteUtils.GetElementMetaData(chapterPageDoc.Root, Consts.Constants.Key_EmbeddedSupplementalModules, xnm);
            if (!string.IsNullOrEmpty(supplementalModulesMetadata))
            {
                var embeddedModulesInfo = EmbeddedModuleInfo.Deserialize(supplementalModulesMetadata);
                var embeddedModuleInfo = embeddedModulesInfo.FirstOrDefault(m => m.ModuleName == moduleInfo.ShortName);
                if (embeddedModuleInfo != null)
                {
                    var tableEl = NotebookGenerator.GetPageTable(chapterPageDoc, xnm);

                    tableEl.XPathSelectElements(string.Format("one:Row/one:Cell[{0}]", embeddedModuleInfo.ColumnIndex + 1), xnm).Remove();
                    tableEl.XPathSelectElements(string.Format("one:Columns/one:Column[{0}]", embeddedModuleInfo.ColumnIndex + 1), xnm).Remove();

                    int index = 0;
                    foreach (var column in tableEl.XPathSelectElements("one:Columns/one:Column", xnm))
                    {
                        column.SetAttributeValue("index", index++);
                    }

                    embeddedModulesInfo.Remove(embeddedModuleInfo);
                    OneNoteUtils.UpdateElementMetaData(chapterPageDoc.Root, 
                        Consts.Constants.Key_EmbeddedSupplementalModules, EmbeddedModuleInfo.Serialize(embeddedModulesInfo), xnm);
                    return true;
                }
            }

            return false;
        }

        public static List<Exception> CheckModules(string primaryModuleName, string parallelModuleName)
        {
            using (var manager = new BibleParallelTranslationManager(primaryModuleName, parallelModuleName, SettingsManager.Instance.NotebookId_Bible))
            {
                manager.ForCheckOnly = true;
                return manager.IterateBaseBible(null, false, true, null, true).Errors;
            }
        }

        public static List<string> CheckForInconsistencies(string baseModuleName, string parallelModuleName)
        {
            using (var manager = new BibleParallelTranslationManager(baseModuleName, parallelModuleName, SettingsManager.Instance.NotebookId_Bible))
            {
                manager.ForCheckOnly = true;
                return manager.IterateBaseBible(null, false, false, null, true).NotFoundBibleVerses;
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
            bool iterateVerses, Action<SimpleVersePointer, SimpleVerse, BibleIteratorArgs> verseAction, bool refreshCache = false)
        {
            _result = new BibleParallelTranslationConnectionResult();

            var bibleVersePointersComparisonTable = BibleParallelTranslationConnectorManager.GetParallelBibleInfo(
                                                          BaseModuleInfo.ShortName, ParallelModuleInfo.ShortName,
                                                          BaseModuleInfo.BibleTranslationDifferences,
                                                          ParallelModuleInfo.BibleTranslationDifferences, refreshCache);            

            foreach (var baseBookContent in BaseBibleInfo.Books)
            {
                var baseBookInfo = BaseModuleInfo.BibleStructure.BibleBooks.FirstOrDefault(b => b.Index == baseBookContent.Index);
                if (baseBookInfo == null)
                    throw new InvalidModuleException(string.Format("Book with index {0} is not found in module manifest", baseBookContent.Index));                

                var parallelBookContent = ParallelBibleInfo.Books.FirstOrDefault(b => b.Index == baseBookContent.Index);
                if (parallelBookContent != null)
                {
                    XElement sectionEl = ForCheckOnly ? null : HierarchySearchManager.FindBibleBookSection(ref _oneNoteApp, BibleNotebookId, baseBookInfo.SectionName);
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
                        _result.Errors.Add(ex);
                    }
                }
                else
                    _result.NotFoundBibleVerses.Add(baseBookInfo.Name);                
            }

            return _result;
        }      

        private void ProcessBibleBook(XElement bibleBookSectionEl, BibleBookInfo baseBookInfo,
            BIBLEBOOK baseBookContent, BIBLEBOOK parallelBookContent, 
            SimpleVersePointersComparisonTable bookVersePointersComparisonTable,
            Func<XDocument, SimpleVersePointer, BibleIteratorArgs> chapterAction, bool needToUpdateChapter,
            bool iterateVerses, Action<SimpleVersePointer, SimpleVerse, BibleIteratorArgs> verseAction)
        {
            XmlNamespaceManager xnm = OneNoteUtils.GetOneNoteXNM();
            string sectionId = ForCheckOnly ? null : (string)bibleBookSectionEl.Attribute("ID");            

            var sectionPagesEl = ForCheckOnly ? null : OneNoteUtils.GetHierarchyElement(ref _oneNoteApp, sectionId, HierarchyScope.hsPages, out xnm);

            int lastProcessedChapter = 0;
            int lastProcessedVerse = 0;            

            foreach (var baseChapter in baseBookContent.Chapters)
            {
                if (Logger != null)                
                    Logger.LogMessage("{0} '{1} {2}'", BibleCommon.Resources.Constants.ProcessChapter, baseBookInfo.Name, baseChapter.Index);

                if (!parallelBookContent.Chapters.Any(pch => pch.Index == baseChapter.Index))
                    _result.NotFoundBibleVerses.Add(string.Format("{0} {1}", baseBookInfo.Name, baseChapter.Index));

                XDocument chapterPageDoc = null;
                BibleIteratorArgs bibleIteratorArgs = null;

                if (chapterAction != null)
                {
                    var chapterPageEl = ForCheckOnly ? null : HierarchySearchManager.FindChapterPage(sectionPagesEl.Root, baseChapter.Index, xnm);

                    if (chapterPageEl == null && !ForCheckOnly)
                        throw new BaseChapterSectionNotFoundException(baseChapter.Index, baseBookInfo.Index);

                    string chapterPageId = ForCheckOnly ? null : (string)chapterPageEl.Attribute("ID");
                    chapterPageDoc = ForCheckOnly ? null : OneNoteUtils.GetPageContent(ref _oneNoteApp, chapterPageId, out xnm);

                    bibleIteratorArgs = chapterAction(chapterPageDoc, new SimpleVersePointer(baseBookInfo.Index, baseChapter.Index));
                }

                var updatingChapterWasNotCanceled = bibleIteratorArgs == null || bibleIteratorArgs.NotNeedToUpdateChapter == null || !bibleIteratorArgs.NotNeedToUpdateChapter.Value;
                var processingVersesWasNotCanceled = bibleIteratorArgs == null || bibleIteratorArgs.NotNeedToProcessVerses == null || !bibleIteratorArgs.NotNeedToProcessVerses.Value;

                bool? chapterWasModified = null;
                if (iterateVerses && processingVersesWasNotCanceled)
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
                                _result.Errors.Add(ex);
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

                if (needToUpdateChapter
                    && chapterAction != null 
                    && chapterWasModified.GetValueOrDefault(true) == true 
                    && !ForCheckOnly 
                    && updatingChapterWasNotCanceled)
                {
                    SupplementalBibleManager.UpdatePageXmlForStrongBible(chapterPageDoc, _isOneNote2010);

                    OneNoteUtils.UpdatePageContentSafe(ref _oneNoteApp, chapterPageDoc, xnm);                    
                }
            }
        }

        private SimpleVerse GetParallelVerse(SimpleVersePointer baseVersePointer, BIBLEBOOK parallelBookContent, 
            SimpleVersePointersComparisonTable bookVersePointersComparisonTable, string strongPrefix, int lastProcessedChapter, int lastProcessedVerse)
        {
            ComparisonVersesInfo parallelVersePointers = new ComparisonVersesInfo();            

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
                    throw new GetParallelVerseException("parallelVersePointers.Count == 0", baseVersePointer, BaseModuleShortName, BaseVersePointerException.Severity.Error);
                
                var parallelVerse = GetParallelVerses(baseVersePointer, parallelVersePointers, parallelBookContent, strongPrefix);
                
                if (!parallelVerse.IsEmpty)
                    CheckVerseForWarnings(baseVersePointer, parallelBookContent, parallelVersePointers.First(), lastProcessedChapter, lastProcessedVerse);  

                return parallelVerse;
            }
            catch (BaseVersePointerException ex)
            {
                if (ex.IsChapterException)
                    throw;

                _result.Errors.Add(ex);
                return new SimpleVerse(baseVersePointer, string.Empty);
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
                                throw new GetParallelVerseException("Miss verse (x01)", baseVersePointer, BaseModuleShortName, BaseVersePointerException.Severity.Warning);
                        }
                        else if (firstParallelVerse.Verse > 1)  // начали главу не с начала                    
                            throw new GetParallelVerseException("Miss verse (x02)", baseVersePointer, BaseModuleShortName, BaseVersePointerException.Severity.Warning);
                    }
                    else
                    {
                        if (lastProcessedVerse > 0 && firstParallelVerse.Verse > lastProcessedVerse + 1)
                            throw new GetParallelVerseException("Miss verse (x03)", baseVersePointer, BaseModuleShortName, BaseVersePointerException.Severity.Warning);
                        else if (lastProcessedChapter == firstParallelVerse.Chapter && lastProcessedVerse == firstParallelVerse.Verse && !firstParallelVerse.PartIndex.HasValue)
                            throw new GetParallelVerseException("Double verse (x04)", baseVersePointer, BaseModuleShortName, BaseVersePointerException.Severity.Warning);
                        else if (lastProcessedChapter == firstParallelVerse.Chapter && firstParallelVerse.Verse < lastProcessedVerse)
                            throw new GetParallelVerseException("Reverse verse (x05)", baseVersePointer, BaseModuleShortName, BaseVersePointerException.Severity.Warning);                        
                    }
                }
            }
            catch (BaseVersePointerException ex)
            {
                _result.Errors.Add(ex);
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
            bool isPartOfBigVerse;


            bool isFullVerses, isDiscontinuous;
            List<SimpleVersePointer> notFoundVerses;
            List<SimpleVersePointer> emptyVerses;
            verseContent = parallelBookContent.GetVersesContent(parallelVersePointers, this.ParallelModuleInfo.ShortName, strongPrefix,
                                        out topLastVerse, out isEmpty, out isFullVerses, out isDiscontinuous, out isPartOfBigVerse, out notFoundVerses, out emptyVerses);

            if (!isEmpty)
            {
                verseNumberContent = GetVersesNumberString(baseVersePointer, parallelVersePointers, topLastVerse, isFullVerses, isDiscontinuous, isPartOfBigVerse, emptyVerses);

                if (verseContent == null)
                {
                    if (!parallelVersePointers.All(pvp => pvp.EmptyVerseContent))
                    {
                        throw new GetParallelVerseException(                                // значит нет такого стиха, либо такой по счёту части стиха      
                            string.Format("Can not find verseContent{0}",
                                            firstParallelVerse.PartIndex.HasValue
                                                ? string.Format(" (versePart = {0})", firstParallelVerse.PartIndex + 1)
                                                : string.Empty),
                                            baseVersePointer, BaseModuleShortName, BaseVersePointerException.Severity.Warning);
                    }
                }
                else
                {
                    foreach (var notFoundVerse in notFoundVerses)
                    {
                        _result.Errors.Add(new GetParallelVerseException(                        // значит один из нескольких стихов не удалось найти
                            string.Format("Can not find verseContent{0}",
                                            notFoundVerse.PartIndex.HasValue
                                                ? string.Format(" (versePart = {0})", notFoundVerse.PartIndex + 1)
                                                : string.Empty),
                                            baseVersePointer, BaseModuleShortName, BaseVersePointerException.Severity.Warning));
                    }
                }
            }

            return new SimpleVerse(firstParallelVerse, verseNumberContent, verseContent)
            {
                VerseNumber = new VerseNumber(firstParallelVerse.Verse, topLastVerse),
                IsEmpty = firstParallelVerse.IsEmpty || isEmpty,
                IsPartOfBigVerse = isPartOfBigVerse
            };
        }

        private string GetVersesNumberString(SimpleVersePointer baseVersePointer, ComparisonVersesInfo parallelVersePointers, 
                                                int? topVerse, bool isFullVerses, bool isDiscontinuous, bool isPartOfBigVerse, List<SimpleVersePointer> emptyVerses)
        {
            string result = string.Empty;
            var notEmptyVerses = parallelVersePointers.Where(v => !v.IsEmpty && !emptyVerses.Contains(v));
            var firstParallelVerse = notEmptyVerses.FirstOrDefault();

            if (firstParallelVerse != null)
            {
                result = GetVerseNumberString(firstParallelVerse, null, baseVersePointer.Chapter, isFullVerses);

                if (notEmptyVerses.Count() > 1 || topVerse.HasValue)
                {
                    var lastVerse = notEmptyVerses.Last();

                    if (lastVerse != firstParallelVerse)
                    {
                        result += string.Format("{0}{1}",
                                                    isDiscontinuous ? ',' : '-',
                                                    GetVerseNumberString(lastVerse, topVerse, firstParallelVerse.Chapter, isFullVerses));
                    }
                    else if (isPartOfBigVerse && topVerse.HasValue)
                        result += "-" + topVerse;

                }
            }

            return result;
        }



        private string GetVerseNumberString(SimpleVersePointer versePointer, int? topVerse, int baseChapter, bool isFullVerses)
        {
            var result = topVerse.HasValue ? topVerse.ToString() : versePointer.VerseNumber.ToString();
            if (baseChapter != versePointer.Chapter)
                result = string.Format("{0}:{1}", versePointer.Chapter, result);

            if (versePointer.PartIndex.HasValue && !isFullVerses)
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

        public static void MergeModuleWithMainBible(ModuleInfo parallelModuleInfo)
        {
            if (!string.IsNullOrEmpty(SettingsManager.Instance.ModuleShortName) 
                && SettingsManager.Instance.ModuleShortName != parallelModuleInfo.ShortName)
            {
                try
                {
                    var baseModuleInfo = ModulesManager.GetModuleInfo(SettingsManager.Instance.ModuleShortName);

                    // merge book abbriviations
                    foreach (var baseBook in baseModuleInfo.BibleStructure.BibleBooks)
                    {
                        var parallelBook = parallelModuleInfo.BibleStructure.BibleBooks.FirstOrDefault(b => b.Index == baseBook.Index);
                        if (parallelBook != null)
                        {
                            foreach (var parallelBookAbbreviation in parallelBook.AllAbbreviations.Values.Where(abbr => string.IsNullOrEmpty(abbr.ModuleName)))
                            {
                                if (!baseBook.AllAbbreviations.ContainsKey(parallelBookAbbreviation.Value))
                                {
                                    baseBook.Abbreviations.Add(new Abbreviation(parallelBookAbbreviation.Value)
                                    {
                                        ModuleName = parallelModuleInfo.ShortName,
                                        IsFullBookName = parallelBookAbbreviation.IsFullBookName
                                    });
                                }
                            }
                        }
                    }

                    //merge alphabets
                    if (!string.IsNullOrEmpty(parallelModuleInfo.BibleStructure.Alphabet))
                    {
                        foreach (var c in parallelModuleInfo.BibleStructure.Alphabet)
                        {
                            if (!baseModuleInfo.BibleStructure.Alphabet.Contains(c))
                                baseModuleInfo.BibleStructure.Alphabet += c;
                        }
                    }

                    ModulesManager.UpdateModuleManifest(baseModuleInfo);
                }
                catch (ModuleNotFoundException) { }
            }
        }       
        
        
        public static void RemoveBookAbbreviationsFromMainBible(string parallelModuleName, bool removeAllParallelModulesAbbriviations)
        {
            if (!string.IsNullOrEmpty(SettingsManager.Instance.ModuleShortName)
                && SettingsManager.Instance.ModuleShortName != parallelModuleName)
            {
                try
                {
                    var baseModuleInfo = ModulesManager.GetModuleInfo(SettingsManager.Instance.ModuleShortName);

                    foreach (var baseBook in baseModuleInfo.BibleStructure.BibleBooks)
                    {
                        baseBook.Abbreviations.RemoveAll(abbr => 
                            (removeAllParallelModulesAbbriviations && !string.IsNullOrEmpty(abbr.ModuleName)) 
                            || (!removeAllParallelModulesAbbriviations && abbr.ModuleName == parallelModuleName));
                    }

                    ModulesManager.UpdateModuleManifest(baseModuleInfo);
                }
                catch (ModuleNotFoundException) { }
            }
        }

        public static void MergeAllModulesWithMainBible()
        {
            foreach (var module in ModulesManager.GetModules(true)
                .Where(m => m.Type == Common.ModuleType.Bible || m.Type == Common.ModuleType.Strong))
            {
                BibleParallelTranslationManager.MergeModuleWithMainBible(module);
            }
        }
    }
}
