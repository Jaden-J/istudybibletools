using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using BibleCommon.Helpers;
using System.Xml.Linq;
using BibleCommon.Consts;
using System.Xml;
using System.Xml.XPath;
using System.Globalization;
using BibleCommon.Common;
using System.Xml.Serialization;
using BibleCommon.Scheme;
using BibleCommon.Contracts;

namespace BibleConfigurator.ModuleConverter
{
    public class ZefaniaXmlConverter: ConverterBase
    {
        public enum ReadParameters
        {
            None,
            //RemoveHyperlinks,
            RemoveStrongs
        }

        protected XMLBIBLE ZefaniaXmlBibleInfo { get; set; }
        protected BibleBooksInfo BooksInfo { get; set; }
        protected string ZefaniaXmlFilePath { get; set; }        

        protected ReadParameters[] AdditionalReadParameters { get; set; }

        public Func<BIBLEBOOK, string, string> ConvertChapterNameFunc { get; set; }
        protected ICustomLogger FormLogger { get; set; }
        

        /// <summary>
        /// 
        /// </summary>
        /// <param name="emptyNotebookName"></param>
        /// <param name="bqModuleFolder"></param>
        /// <param name="manifestFilePathToSave"></param>
        /// <param name="fileEncoding"></param>
        /// <param name="oldTestamentName"></param>
        /// <param name="newTestamentName"></param>
        /// <param name="oldTestamentBooksCount"></param>
        /// <param name="newTestamentBooksCount"></param>
        /// <param name="locale">can be not specified</param>
        /// <param name="notebooksInfo"></param>
        public ZefaniaXmlConverter(string moduleShortName, string moduleName, XMLBIBLE bibleContent, BibleBooksInfo booksInfo, string manifestFilesFolderPath,
            string locale, NotebooksStructure notebooksStructure, BibleTranslationDifferences translationDifferences, 
            string chapterPageNameTemplate, 
            bool isStrong, Version version, Version minProgramVersion, bool generateBibleNotebook, 
            ICustomLogger formLogger,
            params ReadParameters[] readParameters)
            : base(moduleShortName, manifestFilesFolderPath, locale, notebooksStructure, null,
                        translationDifferences, chapterPageNameTemplate, isStrong, 
                        version, minProgramVersion, generateBibleNotebook, true)
        {
            this.ModuleDisplayName = moduleName;            
            this.BooksInfo = booksInfo;
            this.ZefaniaXmlBibleInfo = bibleContent;
            this.BookIndexes = BooksInfo.Books.Where(bi => ZefaniaXmlBibleInfo.Books.Any(zb => zb.Index == bi.Index)).Select(b => b.Index).ToList();

            this.AdditionalReadParameters = readParameters;

            this.FormLogger = formLogger;

            if (this.AdditionalReadParameters == null)
                this.AdditionalReadParameters = new ReadParameters[] { };                
        }

        protected override ExternalModuleInfo ReadExternalModuleInfo()
        {
            return null;
        }        

        protected override void ProcessBibleBooks(ExternalModuleInfo externalModuleInfo)
        {
            XDocument currentChapterDoc = null;
            XElement currentTableElement = null;            
            string currentSectionGroupId = null;

            string oldTestamentName = null;
            int? oldTestamentSectionsCount = null;
            string newTestamentName = null;
            int? newTestamentSectionsCount = null;

            GetTestamentInfo(ContainerType.OldTestament, out oldTestamentName, out oldTestamentSectionsCount);
            GetTestamentInfo(ContainerType.NewTestament, out newTestamentName, out newTestamentSectionsCount);

            this.OldTestamentBooksCount = (oldTestamentSectionsCount ?? newTestamentSectionsCount).Value;


            ZefaniaXmlBibleInfo.Books.ForEach(book =>
            {
                if (!BooksInfo.Books.Any(bi => bi.Index == book.Index))
                    throw new ConverterExceptionBase("There is no book with index {0} in BooksInfo file", book.Index);
            });

            for (int i = 0; i < BooksInfo.Books.Count; i++)             
            {
                var bookInfo = BooksInfo.Books[i];
                var bibleBookContent = ZefaniaXmlBibleInfo.Books.FirstOrDefault(book => book.Index == bookInfo.Index);
                if (bibleBookContent == null)
                {
                    Errors.Add(new ConverterExceptionBase("BibleBook with index '{0}' was not found in ZefaniaXML", bookInfo.Index));
                    continue;
                }

                if (GenerateBibleNotebook)
                    FormLogger.LogMessage(bookInfo.Name);

                var sectionName = GetBookSectionName(bookInfo.Name, BibleInfo.Books.Count);

                if (string.IsNullOrEmpty(currentSectionGroupId))
                    currentSectionGroupId = AddTestamentSectionGroup(oldTestamentName ?? newTestamentName);
                else if (BibleInfo.Books.Count == OldTestamentBooksCount)
                    currentSectionGroupId = AddTestamentSectionGroup(newTestamentName);

                var bookSectionId = AddBookSection(currentSectionGroupId, sectionName, bookInfo.Name);

                foreach (var chapter in bibleBookContent.Chapters)
                {
                    string chapterPageName;
                    if (ConvertChapterNameFunc != null)
                        chapterPageName = ConvertChapterNameFunc(bibleBookContent, chapter.cnumber);
                    else
                        chapterPageName = ConvertChapterName(bookInfo, chapter.cnumber);

                    if (GenerateBibleNotebook)
                        FormLogger.LogMessage(string.Format("{0} {1}", bookInfo.Name, chapter.cnumber));

                    XmlNamespaceManager xnm;
                    currentChapterDoc = AddChapterPage(bookSectionId, chapterPageName, 2, out xnm);
                    currentTableElement = AddTableToChapterPage(currentChapterDoc, xnm);

                    foreach (var verse in chapter.Verses)
                    {
                        try
                        {
                            ProcessVerse(verse, currentTableElement, BooksInfo.Alphabet);
                        }
                        catch (ConverterExceptionBase ex)
                        {
                            Errors.Add(ex);
                        }
                    }
                    
                    UpdateChapterPage(currentChapterDoc);
                }                             
            }           
        }

        protected override void GenerateManifest(ExternalModuleInfo externalModuleInfo)
        {
            ModuleInfo = new ModuleInfo()
            {
                ShortName = ModuleShortName,
                DisplayName = ModuleDisplayName,
                Version = this.Version,
                MinProgramVersion = MinProgramVersion,
                Locale = this.Locale,
                NotebooksStructure = this.NotebooksStructure,
                Type = IsStrong ? BibleCommon.Common.ModuleType.Strong : BibleCommon.Common.ModuleType.Bible
            };
            ModuleInfo.BibleTranslationDifferences = this.TranslationDifferences;
            ModuleInfo.BibleStructure = new BibleStructureInfo()
            {
                Alphabet = BooksInfo.Alphabet,
                ChapterPageNameTemplate = ChapterPageNameTemplate
            };            

            var index = 0;
            foreach (var bookInfo in BibleInfo.Books)
            {
                var bibleBookInfo = BooksInfo.Books.First(b => b.Index == bookInfo.Index);
                var abbriviations = bibleBookInfo.ShortNamesXMLString
                                            .Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
                var shortName = abbriviations.FirstOrDefault(s => s.StartsWith("|"));

                ModuleInfo.BibleStructure.BibleBooks.Add(
                    new BibleBookInfo()
                    {
                        Index = bibleBookInfo.Index,
                        Name = bibleBookInfo.Name,
                        ShortName = !string.IsNullOrEmpty(shortName) ? shortName.Trim(new char[] { '|' }) : null,
                        SectionName = GetBookSectionName(bibleBookInfo.Name, index++),
                        ChapterPageNameTemplate = bibleBookInfo.ChapterPageNameTemplate,
                        Abbreviations = abbriviations.Select(s => 
                                                    new Abbreviation(s.Trim(new char[] { '\'', '|' })) 
                                                    { 
                                                        IsFullBookName = s.StartsWith("'") || s.StartsWith("|")
                                                    }
                                                ).ToList()
                    });
            }

            SaveToXmlFile(ModuleInfo, Constants.ManifestFileName);            
        }

        private string ConvertChapterName(BookInfo bookInfo, string lineText)
        {
            int? chapterIndex = StringUtils.GetStringLastNumber(lineText);
            if (!chapterIndex.HasValue)
                chapterIndex = 1;

            return string.Format(!string.IsNullOrEmpty(bookInfo.ChapterPageNameTemplate)
                                    ? bookInfo.ChapterPageNameTemplate
                                    : this.ChapterPageNameTemplate,
                                 chapterIndex, bookInfo.Name);            
        }

        private void GetTestamentInfo(ContainerType type, out string testamentName, out int? testamentSectionsCount)
        {
            testamentName = null;
            testamentSectionsCount = null;

            var testamentSectionGroup = this.NotebooksStructure.Notebooks.FirstOrDefault(n => n.Type == ContainerType.Bible).SectionGroups.FirstOrDefault(s => s.Type == type);
            if (testamentSectionGroup != null)
            {
                testamentName = testamentSectionGroup.Name;
                testamentSectionsCount = testamentSectionGroup.SectionsCount;
            }
        }

        private void ProcessVerse(VERS verse, XElement currentTableElement, string alphabet)
        {
            if (currentTableElement == null && GenerateBibleNotebook)
                throw new Exception("currentTableElement is null");

            var verseItems = verse.Items;

            if (AdditionalReadParameters.Contains(ReadParameters.RemoveStrongs))            
                verseItems = new object[] { verse.Value };            

            AddVerseRowToTable(currentTableElement, verse.Index, verse.TopIndex, verseItems);
        }       
    }
}
