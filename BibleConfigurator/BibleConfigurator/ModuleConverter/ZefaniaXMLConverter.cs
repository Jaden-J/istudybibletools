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

namespace BibleConfigurator.ModuleConverter
{
    public class ZefaniaXmlConverter: ConverterBase
    {
        public enum ReadParameters
        {
            None,
            RemoveHyperlinks,
            RemoveStrongs
        }

        protected XMLBIBLE ZefaniaXmlBibleInfo { get; set; }
        protected BibleBooksInfo BooksInfo { get; set; }
        protected string ZefaniaXmlFilePath { get; set; }
        protected string ModuleName { get; set; }

        protected ReadParameters[] AdditionalReadParameters { get; set; }

        public Func<BIBLEBOOK, string, string> ConvertChapterNameFunc { get; set; }        
        

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
        public ZefaniaXmlConverter(string moduleShortName, string moduleName, string zefaniaXMLFilePath, BibleBooksInfo booksInfo, string manifestFilesFolderPath, 
            string locale, List<NotebookInfo> notebooksInfo, BibleTranslationDifferences translationDifferences, 
            string chapterSectionNameTemplate, 
            List<SectionInfo> sectionsInfo, bool isStrong, string dictionarySectionGroupName, int? strongNumbersCount,
            Version version, bool generateNotebooks, params ReadParameters[] readParameters)
            : base(moduleShortName, manifestFilesFolderPath, locale, notebooksInfo, null,
                        translationDifferences, chapterSectionNameTemplate, sectionsInfo, isStrong, dictionarySectionGroupName, 
                        strongNumbersCount, version, generateNotebooks, true)
        {
            this.ModuleName = moduleName;
            this.ZefaniaXmlFilePath = zefaniaXMLFilePath;
            this.BooksInfo = booksInfo;
            this.ZefaniaXmlBibleInfo = Utils.LoadFromXmlFile<XMLBIBLE>(ZefaniaXmlFilePath);                
            this.BookIndexes = BooksInfo.Books.Where(bi => ZefaniaXmlBibleInfo.Books.Any(zb => zb.Index == bi.Index)).Select(b => b.Index).ToList();

            this.AdditionalReadParameters = readParameters;            

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

            for (int i = 0; i < BooksInfo.Books.Count; i++)             
            {
                var bookInfo = BooksInfo.Books[i];
                var bibleBookContent = ZefaniaXmlBibleInfo.Books.FirstOrDefault(book => book.Index == bookInfo.Index);
                if (bibleBookContent == null)
                {
                    Errors.Add(new ConverterExceptionBase("BibleBook with index '{0}' was not found in ZefaniaXML", bookInfo.Index));
                    continue;
                }

                var sectionName = GetBookSectionName(bookInfo.Name, BibleInfo.Books.Count);

                if (string.IsNullOrEmpty(currentSectionGroupId))
                    currentSectionGroupId = AddTestamentSectionGroup(oldTestamentName ?? newTestamentName);
                else if (i == OldTestamentBooksCount)
                    currentSectionGroupId = AddTestamentSectionGroup(newTestamentName);

                var bookSectionId = AddBookSection(currentSectionGroupId, sectionName, bookInfo.Name);

                foreach (var chapter in bibleBookContent.Chapters)
                {
                    string chapterPageName;
                    if (ConvertChapterNameFunc != null)
                        chapterPageName = ConvertChapterNameFunc(bibleBookContent, chapter.cnumber);
                    else
                        chapterPageName = ConvertChapterName(bookInfo, chapter.cnumber);

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

        protected override ModuleInfo GenerateManifest(ExternalModuleInfo externalModuleInfo)
        {
            var module = new ModuleInfo()
            {
                ShortName = ModuleShortName,
                Name =  ModuleName,
                Version = this.Version,
                Locale = this.Locale,
                Notebooks = NotebooksInfo,
                Type = IsStrong ? BibleCommon.Common.ModuleType.Strong : BibleCommon.Common.ModuleType.Bible
            };
            module.BibleTranslationDifferences = this.TranslationDifferences;
            module.BibleStructure = new BibleStructureInfo()
            {
                Alphabet = BooksInfo.Alphabet,
                ChapterSectionNameTemplate = ChapterSectionNameTemplate
            };
            module.Sections = this.SectionsInfo;
            module.DictionarySectionGroupName = this.DictionarySectionGroupName;
            module.DictionaryTermsCount = this.StrongNumbersCount;
            
            foreach (var bibleBookInfo in BooksInfo.Books)
            {
                if (BibleInfo.Books.Any(b => b.Index == bibleBookInfo.Index))
                {
                    module.BibleStructure.BibleBooks.Add(
                        new BibleBookInfo()
                        {
                            Index = bibleBookInfo.Index,
                            Name = bibleBookInfo.Name,
                            SectionName = GetBookSectionName(bibleBookInfo.Name, bibleBookInfo.Index),
                            Abbreviations = bibleBookInfo.ShortNamesXMLString.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
                            .Select(s => new Abbreviation(s.Trim(new char[] { '\'' })) { IsFullBookName = s.StartsWith("'") }).ToList()
                        });
                }
            }

            SaveToXmlFile(module, Constants.ManifestFileName);

            return module;
        }

        private string ConvertChapterName(BookInfo bookInfo, string lineText)
        {
            int? chapterIndex = StringUtils.GetStringLastNumber(lineText);
            if (!chapterIndex.HasValue)
                chapterIndex = 1;

            return string.Format(this.ChapterSectionNameTemplate, chapterIndex, bookInfo.Name);            
        }

        private void GetTestamentInfo(ContainerType type, out string testamentName, out int? testamentSectionsCount)
        {
            testamentName = null;
            testamentSectionsCount = null;

            var testamentSectionGroup = this.NotebooksInfo.FirstOrDefault(n => n.Type == ContainerType.Bible).SectionGroups.FirstOrDefault(s => s.Type == type);
            if (testamentSectionGroup != null)
            {
                testamentName = testamentSectionGroup.Name;
                testamentSectionsCount = testamentSectionGroup.SectionsCount;
            }
        }

        private void ProcessVerse(VERS verse, XElement currentTableElement, string alphabet)
        {
            if (currentTableElement == null && GenerateNotebooks)
                throw new Exception("currentTableElement is null");

            var verseItems = verse.Items;

            if (AdditionalReadParameters.Contains(ReadParameters.RemoveStrongs))            
                verseItems = new object[] { verse.Value };            

            AddVerseRowToTable(currentTableElement, verse.Index, verse.TopIndex, verseItems);
        }       
    }
}
