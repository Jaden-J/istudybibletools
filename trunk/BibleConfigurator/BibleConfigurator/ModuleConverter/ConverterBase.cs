using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.OneNote;
using BibleCommon.Helpers;
using BibleCommon.Services;
using System.Xml.Linq;
using BibleCommon.Consts;
using System.Xml;
using System.Xml.XPath;
using BibleCommon.Common;
using System.IO;
using System.Xml.Serialization;
using BibleCommon.Scheme;

namespace BibleConfigurator.ModuleConverter
{
    public class ExternalModuleInfo
    {
        public string Name { get; set; }
        public string ShortName { get; set; }
        public string Alphabet { get; set; }
        public int BooksCount { get; set; }        
    }

    public class ConverterExceptionBase : Exception
    {
        public ConverterExceptionBase(string message, params object[] args)
            :base(string.Format(message, args))
        {
        }
    }

    public class VerseReadException : ConverterExceptionBase
    {
        public VerseReadException(string message, params object[] args)
            : base(message, args)
        {
        }
    }

    public abstract class ConverterBase: IDisposable
    {
        public List<ConverterExceptionBase> Errors { get; set; }

        public ModuleInfo ModuleInfo { get; set; }
        public string BibleNotebookId { get; set; }

        protected abstract ExternalModuleInfo ReadExternalModuleInfo();
        protected abstract void ProcessBibleBooks(ExternalModuleInfo externalModuleInfo);

        protected Application _oneNoteApp;
        protected bool IsStrong { get; set; }
        protected string ModuleShortName { get; set; }        
        protected string ManifestFilesFolderPath { get; set; }                
        protected string Locale { get; set; }
        protected NotebooksStructure NotebooksStructure { get; set; }
        protected XMLBIBLE BibleInfo { get; set; }
        protected BibleTranslationDifferences TranslationDifferences { get; set; }
        protected List<int> BookIndexes { get; set; }  // массив индексов книг. Для KJV - упорядоченный массив цифр от 1 до 66.                 
        protected string ChapterPageNameTemplate { get; set; }        
        protected Version Version { get; set; }
        protected Version MinProgramVersion { get; set; }
        protected int OldTestamentBooksCount { get; set; }
        protected bool GenerateBibleXml { get; set; }        
        protected bool GenerateBibleNotebook { get; set; }
        protected string ModuleDisplayName { get; set; }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="emptyNotebookName"></param>
        /// <param name="manifestFilePathToSave"></param>
        /// <param name="oldTestamentName"></param>
        /// <param name="newTestamentName"></param>
        /// <param name="oldTestamentBooksCount"></param>
        /// <param name="newTestamentBooksCount"></param>
        /// <param name="locale">can be not specified</param>
        /// <param name="notebooksInfo"></param>
        public ConverterBase(string moduleShortName, string manifestFilesFolderPath, 
            string locale, NotebooksStructure notebooksStructure, List<int> bookIndexes, 
            BibleTranslationDifferences translationDifferences, string chapterPageNameTemplate,
            bool isStrong,  
            Version version, Version minProgramVersion, bool generateBibleNotebook, bool generateBibleXml)
        {
            _oneNoteApp = new Application();
            this.IsStrong = isStrong;
            this.ModuleShortName = moduleShortName.ToLower();            
            this.GenerateBibleNotebook = generateBibleNotebook;
            this.GenerateBibleXml = generateBibleXml;
            this.BibleNotebookId = GenerateBibleNotebook ? NotebookGenerator.CreateNotebook(ref _oneNoteApp, ModuleShortName) : null;
            this.ManifestFilesFolderPath = manifestFilesFolderPath;            
            this.Locale = locale;
            this.NotebooksStructure = notebooksStructure;
            this.BibleInfo = new XMLBIBLE();
            this.TranslationDifferences = translationDifferences;            
            this.BookIndexes = bookIndexes;
            this.ChapterPageNameTemplate = chapterPageNameTemplate;                        
            this.Version = version;
            this.MinProgramVersion = minProgramVersion;
            this.Errors = new List<ConverterExceptionBase>();            

            if (!Directory.Exists(ManifestFilesFolderPath))
                Directory.CreateDirectory(ManifestFilesFolderPath);

            CheckModuleParameters();
        }

        private void CheckModuleParameters()
        {
            if (this.IsStrong)
            {
                if (string.IsNullOrEmpty(this.NotebooksStructure.DictionarySectionGroupName))
                    throw new ArgumentNullException("DictionarySectionGroupName");

                if (!this.NotebooksStructure.DictionaryTermsCount.HasValue)
                    throw new ArgumentNullException("StrongNumbersCount");
            }
        }

        public void Convert()
        {
            var externalModuleInfo = ReadExternalModuleInfo();
            
            //UpdateNotebookProperties(externalModuleInfo);            

            ProcessBibleBooks(externalModuleInfo);

            GenerateManifest(externalModuleInfo);

            GenerateBibleInfo(externalModuleInfo);
        }

        protected virtual void GenerateBibleInfo(ExternalModuleInfo externalModuleInfo)
        {
            if (GenerateBibleXml)
            {
                BibleInfo.INFORMATION = new INFORMATION();
                BibleInfo.INFORMATION.Items = new object[] { ModuleDisplayName };
                BibleInfo.INFORMATION.ItemsElementName = new ItemsChoiceType[] { ItemsChoiceType.title };
                SaveToXmlFile(BibleInfo, Constants.BibleContentFileName);             
            }
        }

        protected virtual string GetBookSectionName(string bookName, int bookIndex)
        {
            return NotebookGenerator.GetBibleBookSectionName(bookName, bookIndex, OldTestamentBooksCount);            
        }

        protected virtual void UpdateNotebookProperties(ExternalModuleInfo externalModuleInfo)
        {
            if (!GenerateBibleNotebook)
                return;
            
            XmlNamespaceManager xnm;
            var notebook = OneNoteUtils.GetHierarchyElement(ref _oneNoteApp, BibleNotebookId, HierarchyScope.hsSelf, out xnm);

            string notebookName = Path.GetFileNameWithoutExtension(NotebooksStructure.Notebooks.First(n => n.Type == ContainerType.Bible).Name);

            notebook.Root.SetAttributeValue("name", notebookName);
            notebook.Root.SetAttributeValue("nickname", notebookName);

            OneNoteUtils.UseOneNoteAPI(ref _oneNoteApp, () =>
            {
                _oneNoteApp.UpdateHierarchy(notebook.ToString(), Constants.CurrentOneNoteSchema);
            });
        }

        protected virtual string AddTestamentSectionGroup(string testamentName)
        {
            if (!GenerateBibleNotebook)
                return null;

            return (string)NotebookGenerator.AddRootSectionGroupToNotebook(ref _oneNoteApp, BibleNotebookId, testamentName).Attribute("ID");            
        }

        protected virtual string AddBookSection(string sectionGroupId, string sectionName, string bookName)
        {
            var sectionId = GenerateBibleNotebook ? NotebookGenerator.AddSection(ref _oneNoteApp, sectionGroupId, sectionName) : null;

            if (GenerateBibleXml)
                AddNewBookContent();

            XmlNamespaceManager xnm;
            AddPage(sectionId, bookName, 1, out xnm);               

            return sectionId;
        }

        private void AddNewBookContent()
        {
            int currentBookNumber = BibleInfo.Books.Count;

            if (currentBookNumber > BookIndexes.Count - 1)
                throw new Exception(string.Format("Invalid book indexes: there is no information about book index for book number {0}", currentBookNumber));            

            BibleInfo.BIBLEBOOK = BibleInfo.BIBLEBOOK.Add(new BIBLEBOOK()
            {
                bnumber = BookIndexes[currentBookNumber].ToString()                
            }).ToArray();
        }

        protected virtual void UpdateChapterPage(XDocument chapterPageDoc)
        {
            if (GenerateBibleNotebook)
            {
                OneNoteUtils.UseOneNoteAPI(ref _oneNoteApp, () =>
                {
                    _oneNoteApp.UpdatePageContent(chapterPageDoc.ToString(), DateTime.MinValue, Constants.CurrentOneNoteSchema);
                });
            }
        }

        protected virtual XDocument AddPage(string bookSectionId, string pageTitle, int pageLevel, out XmlNamespaceManager xnm)
        {
            xnm = null;

            if (!GenerateBibleNotebook)           
                return null;            

            return NotebookGenerator.AddPage(ref _oneNoteApp, bookSectionId, pageTitle, pageLevel, Locale, out xnm);    
        }

        protected virtual XDocument AddChapterPage(string bookSectionId, string pageTitle, int pageLevel, out XmlNamespaceManager xnm)
        {
            var pageDoc = AddPage(bookSectionId, pageTitle, pageLevel, out xnm);

            if (GenerateBibleXml)
                AddNewChapterContent();                                     

            return pageDoc;
        }

        private void AddNewChapterContent()
        {
            var currentBook = BibleInfo.Books.Last();
            int currentChapterIndex = currentBook.Chapters.Count + 1;
            currentBook.Items = currentBook.Items.Add(new CHAPTER()
            {
                cnumber = currentChapterIndex.ToString()
            }).ToArray();
        }

        protected virtual XElement AddTableToChapterPage(XDocument chapterDoc, XmlNamespaceManager xnm)
        {
            if (!GenerateBibleNotebook)
                return null;

            return NotebookGenerator.AddTableToPage(chapterDoc, false, xnm, new CellInfo(SettingsManager.Instance.PageWidth_Bible), new CellInfo(NotebookGenerator.MinimalCellWidth));
        }

        protected virtual void AddVerseRowToTable(XElement tableElement, int verseNumber, int? topVerseNumber, object[] verseItems)
        {
            if (GenerateBibleNotebook)
                NotebookGenerator.AddVerseRowToTable(tableElement, BIBLEBOOK.GetFullVerseString(verseNumber, topVerseNumber, VERS.GetVerseText(verseItems)), 1, Locale);

            if (GenerateBibleXml)
            {
                try
                {
                    AddNewVerseContent(verseNumber, topVerseNumber, verseItems);
                }
                catch (ConverterExceptionBase ex)
                {
                    Errors.Add(ex);
                }
            }
        }        

        private void AddNewVerseContent(int verseNumber, int? topVerseNumber, object[] verseItems)
        {
            var currentBook = BibleInfo.Books.Last();
            var currentChapter = currentBook.Chapters.Last();
            var lastVerse = currentChapter.Verses.LastOrDefault();
            int currentVerseIndex = lastVerse != null ? lastVerse.Index + 1 : currentChapter.Verses.Count + 1;            
            if (lastVerse != null && lastVerse.TopIndex.HasValue)
                currentVerseIndex = lastVerse.TopIndex.Value + 1;

            bool throwException = false;
            if (verseNumber != currentVerseIndex)
            {
                for (int i = currentVerseIndex; i <= verseNumber; i++)
                {
                    currentChapter.Items = currentChapter.Items.Add(new VERS()
                    {
                        vnumber = i.ToString(),
                        Items = new object[] { }                        
                    }).ToArray();
                }

                throwException = true;                
            }

            currentChapter.Items = currentChapter.Items.Add(new VERS()
            {
                vnumber = verseNumber.ToString(),
                e = topVerseNumber.ToString(),                
                Items = verseItems
            }).ToArray();

            if (throwException)
                throw new VerseReadException("{0} {1}: expectedVerseIndex != verseIndex: {2} != {3}",
                                                currentBook.Index, currentChapter.Index, currentVerseIndex, verseNumber);
        }

        protected virtual void SaveToXmlFile(object data, string fileName)
        {
            Utils.SaveToXmlFile(data, Path.Combine(ManifestFilesFolderPath, fileName));
        }

        protected virtual void GenerateManifest(ExternalModuleInfo externalModuleInfo)
        {
            var extModuleInfo = (BibleQuotaModuleInfo)externalModuleInfo;

            ModuleInfo = new ModuleInfo()
            {
                ShortName = ModuleShortName,
                DisplayName = extModuleInfo.Name,
                Version = this.Version,
                MinProgramVersion = MinProgramVersion,
                Locale = this.Locale,
                NotebooksStructure = this.NotebooksStructure,
                Type = IsStrong ? BibleCommon.Common.ModuleType.Strong : BibleCommon.Common.ModuleType.Bible
            };
            ModuleInfo.BibleTranslationDifferences = this.TranslationDifferences;
            ModuleInfo.BibleStructure = new BibleStructureInfo()
            {
                Alphabet = extModuleInfo.Alphabet,                
                ChapterPageNameTemplate = ChapterPageNameTemplate              
            };            

            int index = 0;
            foreach (var bibleBookInfo in extModuleInfo.BibleBooksInfo)
            {
                ModuleInfo.BibleStructure.BibleBooks.Add(
                    new BibleBookInfo()
                    {
                        Index = BookIndexes[index++],
                        Name = bibleBookInfo.Name,
                        SectionName = bibleBookInfo.SectionName,
                        Abbreviations = bibleBookInfo.Abbreviations
                    });
            }

            SaveToXmlFile(ModuleInfo, Constants.ManifestFileName);            
        }

        public void Dispose()
        {
            _oneNoteApp = null;
        }
    }
}
