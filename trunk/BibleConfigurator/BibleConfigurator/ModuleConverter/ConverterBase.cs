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

    public abstract class ConverterBase
    {
        protected abstract ExternalModuleInfo ReadExternalModuleInfo();
        protected abstract void ProcessBibleBooks(ExternalModuleInfo externalModuleInfo);        

        protected Application OneNoteApp { get; set; }
        protected string EmptyNotebookName { get; set; }
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
        public ConverterBase(string emptyNotebookName, string manifestFilesFolderPath,
            string oldTestamentName, string newTestamentName, int oldTestamentBooksCount, int newTestamentBooksCount,
            string locale, List<NotebookInfo> notebooksInfo, List<int> bookIndexes, BibleTranslationDifferences translationDifferences)
        {
            OneNoteApp = new Application();
            this.EmptyNotebookName = emptyNotebookName;
            this.NotebookId = OneNoteUtils.GetNotebookIdByName(OneNoteApp, EmptyNotebookName, true);
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
            int bookPrefix = bookIndex + 1 > OldTestamentBooksCount ? bookIndex + 1 - OldTestamentBooksCount : bookIndex + 1;
            return string.Format("{0:00}. {1}", bookPrefix, bookName);
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
            return OneNoteUtils.AddRootSectionGroupToNotebook(OneNoteApp, NotebookId, testamentName).Attribute("ID").Value;            
        }

        protected virtual string AddBookSection(string sectionGroupId, string sectionName, string bookName)
        {
            var sectionId = NotebookGenerator.AddBookSectionToBibleNotebook(OneNoteApp, sectionGroupId, sectionName, bookName);

            AddNewBookContent();

            XmlNamespaceManager xnm;
            //AddChapterPage(sectionId, bookName, 1, out xnm);               

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

        protected virtual XDocument AddChapterPage(string bookSectionId, string pageTitle, int pageLevel, out XmlNamespaceManager xnm)
        {
            var pageDoc = NotebookGenerator.AddChapterPageToBibleNotebook(OneNoteApp, bookSectionId, pageTitle, pageLevel, Locale, out xnm);    

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
            return NotebookGenerator.AddTableToBibleChapterPage(chapterDoc, SettingsManager.Instance.PageWidth_Bible, xnm);
        }

        protected virtual void AddVerseRowToTable(XElement tableElement, string verseText)
        {
            NotebookGenerator.AddVerseRowToBibleTable(tableElement, verseText, Locale);            

            AddNewVerseContent(verseText);            
        }

        private void AddNewVerseContent(string verseText)
        {
            var currentBook = BibleInfo.Content.Books.Last();
            var currentChapter = currentBook.Chapters.Last();
            int currentVerseIndex = currentChapter.Verses.Count + 1;

            int? verseIndex = StringUtils.GetStringFirstNumber(verseText);
            if (verseIndex.GetValueOrDefault(0) != currentVerseIndex)
                throw new InvalidDataException(string.Format("verseIndex != currentVerseIndex: {0} != {1}", verseIndex, currentVerseIndex));

            currentChapter.Verses.Add(new BibleVerseContent()
            {
                Index = currentVerseIndex,
                Value = verseText
            });
        }

        protected virtual void SaveToXmlFile(object data, string fileName)
        {
            XmlSerializer ser = new XmlSerializer(data.GetType());
            using (var fs = new FileStream(Path.Combine(ManifestFilesFolderPath, fileName), FileMode.Create))
            {
                ser.Serialize(fs, data);
                fs.Flush();
            }
        }

        protected virtual void GenerateManifest(ExternalModuleInfo externalModuleInfo)
        {
            var extModuleInfo = (BibleQuotaModuleInfo)externalModuleInfo;

            ModuleInfo module = new ModuleInfo() { Name = extModuleInfo.Name, Version = "1.0", Notebooks = NotebooksInfo };
            module.BibleTranslationDifferences = this.TranslationDifferences;
            module.BibleStructure = new BibleStructureInfo()
            {
                Alphabet = extModuleInfo.Alphabet,
                NewTestamentName = NewTestamentName,
                OldTestamentName = OldTestamentName,
                OldTestamentBooksCount = OldTestamentBooksCount,
                NewTestamentBooksCount = NewTestamentBooksCount                                
            };

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
    }
}
