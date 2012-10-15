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
    public class BibleQuotaModuleInfo : ExternalModuleInfo
    {
        public string ChapterSign { get; set; }
        public string VerseSign { get; set; }

        public List<BibleQuotaBibleBookInfo> BibleBooksInfo { get; set; } 

        public BibleQuotaModuleInfo()
        {
            BibleBooksInfo = new List<BibleQuotaBibleBookInfo>();
        }
    }

    public class BibleQuotaBibleBookInfo 
    {
        public string Name { get; set; }
        public List<Abbreviation> Abbreviations { get; set; }
        public int ChaptersCount { get; set; }
        public string FileName { get; set; }
        public string SectionName { get; set; }
    }    

    public class BibleQuotaConverter: ConverterBase
    {
        public enum ReadParameters
        {
            None,
            RemoveHyperlinks,
            RemoveStrongs
        }

        protected const string IniFileName = "bibleqt.ini";
        protected string ModuleFolder { get; set; }
        protected ReadParameters[] AdditionalReadParameters { get; set; }

        public Func<BibleQuotaBibleBookInfo, string, string> ConvertChapterNameFunc { get; set; }        
        

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
        public BibleQuotaConverter(string moduleShortName, string bqModuleFolder, string manifestFilesFolderPath, 
            string locale, List<NotebookInfo> notebooksInfo, List<int> bookIndexes, BibleTranslationDifferences translationDifferences, 
            string chapterSectionNameTemplate, List<SectionInfo> sectionsInfo,
            bool isStrong, string dictionarySectionGroupName, int? strongNumbersCount,
            string version, bool generateNotebooks, params ReadParameters[] readParameters)
            : base(moduleShortName, manifestFilesFolderPath, locale, notebooksInfo, bookIndexes,
                        translationDifferences, chapterSectionNameTemplate, sectionsInfo, isStrong, dictionarySectionGroupName, 
                        strongNumbersCount, version, generateNotebooks, true)
        {
            this.ModuleFolder = bqModuleFolder;
            this.AdditionalReadParameters = readParameters;

            if (this.AdditionalReadParameters == null)
                this.AdditionalReadParameters = new ReadParameters[] { };                
        }

        protected override ExternalModuleInfo ReadExternalModuleInfo()
        {
            string iniFilePath = Path.Combine(ModuleFolder, IniFileName);            

            var result = new BibleQuotaModuleInfo();

            foreach (string line in File.ReadAllLines(iniFilePath, Utils.GetFileEncoding(iniFilePath)))
            {
                var pair = line.Split(new char[] { '=' }, 2, StringSplitOptions.RemoveEmptyEntries);
                if (pair.Length == 2)
                {
                    string key = pair[0].Trim();
                    string value = pair[1].Trim();

                    if (key == "BibleName")
                        result.Name = value;
                    else if (key == "BibleShortName")
                        result.ShortName = value;
                    else if (key == "Alphabet")
                        result.Alphabet = value;
                    else if (key == "BookQty")
                        result.BooksCount = int.Parse(value);
                    else if (key == "ChapterSign")
                        result.ChapterSign = value;
                    else if (key == "VerseSign")
                        result.VerseSign = value;
                    else if (key == "PathName")
                        result.BibleBooksInfo.Add(new BibleQuotaBibleBookInfo() { FileName = value });
                    else if (key == "FullName")
                        result.BibleBooksInfo[result.BibleBooksInfo.Count - 1].Name = value;
                    else if (key == "ShortName")
                        result.BibleBooksInfo[result.BibleBooksInfo.Count - 1].Abbreviations 
                            = value
                                .ToLowerInvariant()
                                .Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)
                                .Select(s => s.Trim(new char[] { '.' })).Distinct()
                                .Where(s => s.All(c => StringUtils.IsDigit(c) || char.IsSymbol(c) || StringUtils.IsCharAlphabetical(c, result.Alphabet, true)))  // чтобы отсечь сокращения на других языках. Потому что на другиз языках другие модули
                                .Select(s => new Abbreviation(s)).ToList();
                    else if (key == "ChapterQty")
                        result.BibleBooksInfo[result.BibleBooksInfo.Count - 1].ChaptersCount = int.Parse(value);
                }
            }

            return result;
        }

        protected override void ProcessBibleBooks(ExternalModuleInfo externalModuleInfo)
        {
            var moduleInfo = (BibleQuotaModuleInfo)externalModuleInfo;

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

            for (int i = 0; i < moduleInfo.BibleBooksInfo.Count; i++) 
            {
                var bibleBookInfo = moduleInfo.BibleBooksInfo[i];                
                bibleBookInfo.SectionName = GetBookSectionName(bibleBookInfo.Name, i);

                if (string.IsNullOrEmpty(currentSectionGroupId))
                    currentSectionGroupId = AddTestamentSectionGroup(oldTestamentName ?? newTestamentName);
                else if (i == OldTestamentBooksCount)
                    currentSectionGroupId = AddTestamentSectionGroup(newTestamentName);

                var bookSectionId = AddBookSection(currentSectionGroupId, bibleBookInfo.SectionName, bibleBookInfo.Name);

                string bookFile = Path.Combine(ModuleFolder, bibleBookInfo.FileName);                

                foreach (string line in File.ReadAllLines(bookFile, Utils.GetFileEncoding(bookFile)))
                {
                    string lineText = ShellText(line, moduleInfo); 

                    if (line.StartsWith(moduleInfo.ChapterSign))
                    {                        
                        if (currentChapterDoc != null)
                            UpdateChapterPage(currentChapterDoc);

                        if (ConvertChapterNameFunc != null)
                            lineText = ConvertChapterNameFunc(bibleBookInfo, lineText);
                        else                       
                            lineText = ConvertChapterName(bibleBookInfo, lineText);                                                   

                        XmlNamespaceManager xnm;
                        currentChapterDoc = AddChapterPage(bookSectionId, lineText, 2, out xnm);

                        currentTableElement = AddTableToChapterPage(currentChapterDoc, xnm);
                    }
                    else if (line.StartsWith(moduleInfo.VerseSign))
                    {
                        try
                        {
                            ProcessVerse(lineText, currentTableElement, externalModuleInfo.Alphabet);
                        }
                        catch (ConverterExceptionBase ex)
                        {
                            Errors.Add(ex);
                        }
                    }
                }
            }

            if (currentChapterDoc != null)
            {
                UpdateChapterPage(currentChapterDoc);
            }
        }    

        private string ConvertChapterName(BibleQuotaBibleBookInfo bibleBookInfo, string lineText)
        {
            int? chapterIndex = StringUtils.GetStringLastNumber(lineText);
            if (!chapterIndex.HasValue)
                chapterIndex = 1;
            
            return string.Format(this.ChapterSectionNameTemplate, chapterIndex, bibleBookInfo.Name);            
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

        private void ProcessVerse(string lineText, XElement currentTableElement, string alphabet)
        {
            if (currentTableElement == null && GenerateNotebooks)
                throw new Exception("currentTableElement is null");

            int? verseNumber;
            int? topVerseNumber;
            if (!string.IsNullOrEmpty(lineText))
            {
                string verseText = GetVerseTextWithoutNumber(lineText, out verseNumber, out topVerseNumber);

                if (!verseNumber.HasValue)
                {
                    var currentBook = BibleInfo.Books.Last();
                    var currentChapter = currentBook.Chapters.Last();
                    throw new VerseReadException("{0} {1}: Verse has no number: {2}", currentBook.Index, currentChapter.Index, lineText);
                }

                 вот здесь. надо сконвертить ZEFANIA XML модуль со стронгом. и так же научиться его считывать (создавать спр Библию), учитывая перфиксы (что их теперь надо подставлять во время создания спр Библии) и т.д.

                List<object> verseItems = null;
                if (IsStrong || AdditionalReadParameters.Contains(ReadParameters.RemoveStrongs))
                {
                    verseItems = GetStrongVerseItems(verseText, alphabet);                    
                    verseItems = ProcessStrongVerse(verseItems);
                }
                else
                    verseItems = new List<object>() { verseText };

                AddVerseRowToTable(currentTableElement, verseNumber.Value, topVerseNumber, verseItems.ToArray());
            }
        }

        private List<object> ProcessStrongVerse(List<object> verseItems)
        {
            var result = new List<object>();

            if (!AdditionalReadParameters.Contains(ReadParameters.RemoveStrongs))
            {
                foreach (var verseItem in verseItems)
                {
                    var prev = result.LastOrDefault();

                    if (prev != null)
                    {                        
                        if (verseItem is string)
                        {
                            if (prev is GRAM && string.IsNullOrEmpty((string)verseItem))
                                ((GRAM)prev).str += " ";
                            else
                                result.Add(verseItem);
                        }
                        else if (verseItem is GRAM)
                        {
                            if (prev is GRAM)
                                ((GRAM)prev).str += ((GRAM)verseItem).str;
                            else
                            {
                                ((GRAM)verseItem).Items = new object[] { prev };
                                result[result.Count - 1] = verseItem;                                
                            }
                        }
                    }
                    else
                        result.Add(verseItem);
                }
                verseItems = result;
            }           

            return result;
        }

        protected override void GenerateBibleInfo(ModuleInfo moduleInfo)
        {
            base.GenerateBibleInfo(moduleInfo);

            var booksInfo = new BibleBooksInfo() { Descr = this.ModuleShortName, Alphabet = moduleInfo.BibleStructure.Alphabet };
            foreach (var book in moduleInfo.BibleStructure.BibleBooks)
            {
                booksInfo.Books.Add(new BookInfo()
                {
                    Index = book.Index,
                    Name = book.Name,
                    ShortNamesXMLString = string.Join(";", book.Abbreviations.ConvertAll(a => a.Value).ToArray())
                });
            }
            SaveToXmlFile(booksInfo, "BibleBooksInfo.xml");
        }

        private List<object> GetStrongVerseItems(string verseText, string alphabet)
        {
            var result = new List<object>();
            int currentBookNumber = BibleInfo.Books.Count;            

            int cursorPosition = StringUtils.GetNextIndexOfDigit(verseText, null);
            if (cursorPosition > -1)
            {
                int textBreakIndex, htmlBreakIndex = -1;
                string strongNumber = StringUtils.GetNextString(verseText, cursorPosition - 1, new SearchMissInfo(0, SearchMissInfo.MissMode.CancelOnMissFound), alphabet,
                                                                    out textBreakIndex, out htmlBreakIndex, StringSearchIgnorance.None, StringSearchMode.SearchNumber);
                if (!string.IsNullOrEmpty(strongNumber))
                {
                    var text = verseText.Substring(0, cursorPosition);
                    if (!string.IsNullOrEmpty(text))
                        result.Add(text.TrimEnd());

                    if (AdditionalReadParameters.Contains(ReadParameters.RemoveStrongs))
                        cursorPosition -= 1;  // чтобы удалить пробел перед номером стронга                    
                    else
                        result.Add(new GRAM() { str = strongNumber });

                    result.AddRange(GetStrongVerseItems(verseText.Substring(htmlBreakIndex), alphabet));
                }
            }
            else
                result.Add(verseText);

            return result;
        }

        private string GetVerseTextWithoutNumber(string lineText, out int? verseNumber, out int? topVerseNumber)
        {
            verseNumber = null;
            topVerseNumber = null;

            if (!string.IsNullOrEmpty(lineText.Trim()))
            {
                if (StringUtils.IsDigit(lineText[0]))
                {
                    verseNumber = StringUtils.GetStringFirstNumber(lineText);

                    lineText = lineText.Remove(0, verseNumber.Value.ToString().Length).Trim();

                    if (lineText.Length >= 2)
                    {
                        if (lineText[0] == '-' && StringUtils.IsDigit(lineText[1]))
                        {
                            topVerseNumber = StringUtils.GetStringFirstNumber(lineText);
                            lineText = lineText.Remove(0, topVerseNumber.Value.ToString().Length + 1).Trim();
                        }
                    }
                }
            }

            return lineText;
        }

        private string ShellText(string line, BibleQuotaModuleInfo moduleInfo)
        {
            var result = line.Replace("<<", "&lt;").Replace(">>", "&gt;");   // чтобы учитывать строки типа "<p>1 <<To the chief Musician on Neginoth, A Psalm of David.>> Hear me when I call, O God of my righteousness: thou hast enlarged me when I was in distress; have mercy upon me, and hear my prayer."
            result = StringUtils.RemoveIllegalTagStartAndEndSymbols(result);

            if (AdditionalReadParameters.Contains(ReadParameters.RemoveHyperlinks))
            {
                result = StringUtils.RemoveTags(result, "<a>", "</a>");
                result = StringUtils.RemoveTags(result, "<a ", "</a>");                
            }

            result = result.Replace("  ", " ");

            result = StringUtils.GetText(result, moduleInfo.Alphabet).Trim();            

            return result;
        }
    }
}
