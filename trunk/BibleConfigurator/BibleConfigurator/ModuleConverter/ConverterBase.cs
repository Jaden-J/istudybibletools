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

namespace BibleConfigurator.ModuleConverter
{
    public class ExternalModuleInfo
    {
        public string Name { get; set; }
        public string ShortName { get; set; }
        public string Alphabet { get; set; }
        public int BooksCount { get; set; }        
    }

    public abstract class ConverterExceptionBase : Exception
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

        protected abstract ExternalModuleInfo ReadExternalModuleInfo();
        protected abstract void ProcessBibleBooks(ExternalModuleInfo externalModuleInfo);        

        protected Application OneNoteApp { get; set; }
        protected bool IsStrong { get; set; }
        protected string NewNotebookName { get; set; }
        protected string NotebookId { get; set; }
        protected string ManifestFilesFolderPath { get; set; }        
        protected string OldTestamentName { get; set; }
        protected string NewTestamentName { get; set; }        
        protected int OldTestamentBooksCount { get; set; }
        protected int NewTestamentBooksCount { get; set; }
        protected string Locale { get; set; }
        protected List<NotebookInfo> NotebooksInfo { get; set; }
        protected ModuleBibleInfo BibleInfo { get; set; }
        protected BibleTranslationDifferences TranslationDifferences { get; set; }
        protected List<int> BookIndexes { get; set; }  // массив индексов книг. Для KJV - упорядоченный массив цифр от 1 до 66.                 
        protected string ChapterSectionNameTemplate { get; set; }
        protected List<SectionInfo> SectionsInfo { get; set; }
        protected string DictionarySectionGroupName { get; set; }
        protected string Version { get; set; }        

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
        public ConverterBase(string newNotebookName, string manifestFilesFolderPath, bool isStrong,
            string oldTestamentName, string newTestamentName, int oldTestamentBooksCount, int newTestamentBooksCount,
            string locale, List<NotebookInfo> notebooksInfo, List<int> bookIndexes, 
            BibleTranslationDifferences translationDifferences, string chapterSectionNameTemplate,
            List<SectionInfo> sectionsInfo, string dictionarySectionGroupName, string version)
        {
            OneNoteApp = new Application();
            this.IsStrong = isStrong;
            this.NewNotebookName = newNotebookName;
            this.NotebookId = NotebookGenerator.CreateNotebook(OneNoteApp, NewNotebookName);
            this.ManifestFilesFolderPath = manifestFilesFolderPath;
            this.OldTestamentName = oldTestamentName;
            this.NewTestamentName = newTestamentName;
            this.OldTestamentBooksCount = oldTestamentBooksCount;
            this.NewTestamentBooksCount = newTestamentBooksCount;
            this.Locale = locale;
            this.NotebooksInfo = notebooksInfo;
            this.BibleInfo = new ModuleBibleInfo();
            this.TranslationDifferences = translationDifferences;
            this.BibleInfo.Content.Locale = locale;
            this.BookIndexes = bookIndexes;
            this.ChapterSectionNameTemplate = chapterSectionNameTemplate;
            this.SectionsInfo = sectionsInfo;
            this.DictionarySectionGroupName = dictionarySectionGroupName;
            this.Version = version;
            this.Errors = new List<ConverterExceptionBase>();

            if (!Directory.Exists(manifestFilesFolderPath))
                Directory.CreateDirectory(manifestFilesFolderPath);
        }

        public void Convert()
        {
            var externalModuleInfo = ReadExternalModuleInfo();
            
            //UpdateNotebookProperties(externalModuleInfo);            

            ProcessBibleBooks(externalModuleInfo);

            GenerateManifest(externalModuleInfo);

            GenerateBibleInfo();
        }

        protected virtual void GenerateBibleInfo()
        {
            SaveToXmlFile(BibleInfo, Constants.BibleInfoFileName);
        }

        protected virtual string GetBookSectionName(string bookName, int bookIndex)
        {
            return NotebookGenerator.GetBibleBookSectionName(bookName, bookIndex, OldTestamentBooksCount);            
        }

        protected virtual void UpdateNotebookProperties(ExternalModuleInfo externalModuleInfo)
        {
            XmlNamespaceManager xnm;
            var notebook = OneNoteUtils.GetHierarchyElement(OneNoteApp, NotebookId, HierarchyScope.hsSelf, out xnm);

            string notebookName = Path.GetFileNameWithoutExtension(NotebooksInfo.First(n => n.Type == NotebookType.Bible).Name);

            notebook.Root.SetAttributeValue("name", notebookName);
            notebook.Root.SetAttributeValue("nickname", notebookName);

            OneNoteApp.UpdateHierarchy(notebook.ToString(), Constants.CurrentOneNoteSchema);
        }

        protected virtual string AddTestamentSectionGroup(string testamentName)
        {
            return NotebookGenerator.AddRootSectionGroupToNotebook(OneNoteApp, NotebookId, testamentName).Attribute("ID").Value;            
        }

        protected virtual string AddBookSection(string sectionGroupId, string sectionName, string bookName)
        {
            var sectionId = NotebookGenerator.AddSection(OneNoteApp, sectionGroupId, sectionName);

            AddNewBookContent();

            XmlNamespaceManager xnm;
            AddPage(sectionId, bookName, 1, out xnm);               

            return sectionId;
        }

        private void AddNewBookContent()
        {
            int currentBookNumber = BibleInfo.Content.Books.Count;
            BibleInfo.Content.Books.Add(new BibleBookContent()
            {
                Index = BookIndexes[currentBookNumber]
            });
        }

        protected virtual void UpdateChapterPage(XDocument chapterPageDoc)
        {            
            OneNoteApp.UpdatePageContent(chapterPageDoc.ToString(), DateTime.MinValue, Constants.CurrentOneNoteSchema);                                 
        }

        protected virtual XDocument AddPage(string bookSectionId, string pageTitle, int pageLevel, out XmlNamespaceManager xnm)
        {
            return NotebookGenerator.AddPage(OneNoteApp, bookSectionId, pageTitle, pageLevel, Locale, out xnm);    
        }

        protected virtual XDocument AddChapterPage(string bookSectionId, string pageTitle, int pageLevel, out XmlNamespaceManager xnm)
        {
            var pageDoc = AddPage(bookSectionId, pageTitle, pageLevel, out xnm);

            AddNewChapterContent();                                     

            return pageDoc;
        }

        private void AddNewChapterContent()
        {
            var currentBook = BibleInfo.Content.Books.Last();
            int currentChapterIndex = currentBook.Chapters.Count + 1;
            currentBook.Chapters.Add(new BibleChapterContent()
            {
                Index = currentChapterIndex
            });
        }

        protected virtual XElement AddTableToChapterPage(XDocument chapterDoc, XmlNamespaceManager xnm)
        {
            return NotebookGenerator.AddTableToPage(chapterDoc, false, xnm, new CellInfo(SettingsManager.Instance.PageWidth_Bible), new CellInfo(NotebookGenerator.MinimalCellWidth));
        }

        protected virtual void AddVerseRowToTable(XElement tableElement, int verseNumber, string verseText)
        {
            NotebookGenerator.AddVerseRowToTable(tableElement, string.Format("{0} {1}", verseNumber, verseText), 1, Locale);

            try
            {
                AddNewVerseContent(verseNumber, verseText);
            }
            catch (ConverterExceptionBase ex)
            {
                Errors.Add(ex);
            }
        }

        private void AddNewVerseContent(int verseNumber, string verseText)
        {
            var currentBook = BibleInfo.Content.Books.Last();
            var currentChapter = currentBook.Chapters.Last();
            int currentVerseIndex = currentChapter.Verses.Count + 1;

            if (verseNumber != currentVerseIndex)
            {
                for (int i = currentVerseIndex; i <= verseNumber; i++)
                {
                    currentChapter.Verses.Add(new BibleVerseContent()
                    {
                        Index = i, Value = string.Empty
                    });
                }

                throw new VerseReadException("{0} {1}: expectedVerseIndex != verseIndex: {2} != {3}", 
                                                currentBook.Index, currentChapter.Index, currentVerseIndex, verseNumber);
            }

            currentChapter.Verses.Add(new BibleVerseContent()
            {
                Index = currentVerseIndex,
                Value = verseText
            });
        }

        protected virtual void SaveToXmlFile(object data, string fileName)
        {
            Utils.SaveToXmlFile(data, Path.Combine(ManifestFilesFolderPath, fileName));
        }

        protected virtual void GenerateManifest(ExternalModuleInfo externalModuleInfo)
        {
            var extModuleInfo = (BibleQuotaModuleInfo)externalModuleInfo;

            ModuleInfo module = new ModuleInfo() 
            { 
                Name = extModuleInfo.Name, 
                Version = this.Version, 
                Notebooks = NotebooksInfo, 
                Type = IsStrong ? ModuleType.Strong: ModuleType.Bible 
            };
            module.BibleTranslationDifferences = this.TranslationDifferences;
            module.BibleStructure = new BibleStructureInfo()
            {
                Alphabet = extModuleInfo.Alphabet,
                NewTestamentName = NewTestamentName,
                OldTestamentName = OldTestamentName,
                OldTestamentBooksCount = OldTestamentBooksCount,
                NewTestamentBooksCount = NewTestamentBooksCount,
                ChapterSectionNameTemplate = ChapterSectionNameTemplate              
            };
            module.Sections = this.SectionsInfo;
            module.DictionarySectionGroupName = this.DictionarySectionGroupName;

            int index = 0;
            foreach (var bibleBookInfo in extModuleInfo.BibleBooksInfo)
            {
                module.BibleStructure.BibleBooks.Add(
                    new BibleBookInfo()
                    {
                        Index = BookIndexes[index++],
                        Name = bibleBookInfo.Name,
                        SectionName = bibleBookInfo.SectionName,
                        Abbreviations = bibleBookInfo.Abbreviations
                    });
            }

            SaveToXmlFile(module, Constants.ManifestFileName);
        }

        public void Dispose()
        {
            OneNoteApp = null;
        }
    }
}
