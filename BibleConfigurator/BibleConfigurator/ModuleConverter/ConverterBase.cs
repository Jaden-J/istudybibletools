﻿using System;
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
        protected string ModuleShortName { get; set; }
        protected string NotebookId { get; set; }
        protected string ManifestFilesFolderPath { get; set; }                
        protected string Locale { get; set; }
        protected List<NotebookInfo> NotebooksInfo { get; set; }
        protected ModuleBibleInfo BibleInfo { get; set; }
        protected BibleTranslationDifferences TranslationDifferences { get; set; }
        protected List<int> BookIndexes { get; set; }  // массив индексов книг. Для KJV - упорядоченный массив цифр от 1 до 66.                 
        protected string ChapterSectionNameTemplate { get; set; }
        protected List<SectionInfo> SectionsInfo { get; set; }
        protected string DictionarySectionGroupName { get; set; }
        public int? StrongNumbersCount { get; set; }
        protected string Version { get; set; }
        protected int OldTestamentBooksCount { get; set; }
        protected bool GenerateXmlOnly { get; set; }        

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
            string locale, List<NotebookInfo> notebooksInfo, List<int> bookIndexes, 
            BibleTranslationDifferences translationDifferences, string chapterSectionNameTemplate, List<SectionInfo> sectionsInfo,
            bool isStrong, string dictionarySectionGroupName, int? strongNumbersCount, 
            string version, bool generateXmlOnly)
        {
            OneNoteApp = new Application();
            this.IsStrong = isStrong;
            this.ModuleShortName = moduleShortName;            
            this.GenerateXmlOnly = generateXmlOnly;            
            this.NotebookId = !GenerateXmlOnly ? NotebookGenerator.CreateNotebook(OneNoteApp, ModuleShortName) : null;
            this.ManifestFilesFolderPath = manifestFilesFolderPath;            
            this.Locale = locale;
            this.NotebooksInfo = notebooksInfo;
            this.BibleInfo = new ModuleBibleInfo();
            this.TranslationDifferences = translationDifferences;
            this.BibleInfo.Content.Locale = locale;
            this.BookIndexes = bookIndexes;
            this.ChapterSectionNameTemplate = chapterSectionNameTemplate;
            this.SectionsInfo = sectionsInfo;
            this.DictionarySectionGroupName = dictionarySectionGroupName;
            this.StrongNumbersCount = strongNumbersCount;
            this.Version = version;
            this.Errors = new List<ConverterExceptionBase>();            


            if (!Directory.Exists(ManifestFilesFolderPath))
                Directory.CreateDirectory(ManifestFilesFolderPath);

            CheckModuleParameters();
        }

        private void CheckModuleParameters()
        {
            if (this.IsStrong)
            {
                if (string.IsNullOrEmpty(this.DictionarySectionGroupName))
                    throw new ArgumentNullException("DictionarySectionGroupName");

                if (!this.StrongNumbersCount.HasValue)
                    throw new ArgumentNullException("StrongNumbersCount");
            }
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
            if (GenerateXmlOnly)
                return;
            
            XmlNamespaceManager xnm;
            var notebook = OneNoteUtils.GetHierarchyElement(OneNoteApp, NotebookId, HierarchyScope.hsSelf, out xnm);

            string notebookName = Path.GetFileNameWithoutExtension(NotebooksInfo.First(n => n.Type == ContainerType.Bible).Name);

            notebook.Root.SetAttributeValue("name", notebookName);
            notebook.Root.SetAttributeValue("nickname", notebookName);

            OneNoteApp.UpdateHierarchy(notebook.ToString(), Constants.CurrentOneNoteSchema);
        }

        protected virtual string AddTestamentSectionGroup(string testamentName)
        {
            if (GenerateXmlOnly)
                return null;

            return (string)NotebookGenerator.AddRootSectionGroupToNotebook(OneNoteApp, NotebookId, testamentName).Attribute("ID");            
        }

        protected virtual string AddBookSection(string sectionGroupId, string sectionName, string bookName)
        {
            var sectionId = !GenerateXmlOnly ? NotebookGenerator.AddSection(OneNoteApp, sectionGroupId, sectionName) : null;

            AddNewBookContent();

            XmlNamespaceManager xnm;
            AddPage(sectionId, bookName, 1, out xnm);               

            return sectionId;
        }

        private void AddNewBookContent()
        {
            int currentBookNumber = BibleInfo.Content.Books.Count;

            if (currentBookNumber > BookIndexes.Count - 1)
                throw new Exception(string.Format("Invalid book indexes: there is no information about book index for book number {0}", currentBookNumber));

            BibleInfo.Content.Books.Add(new BibleBookContent()
            {
                Index = BookIndexes[currentBookNumber]
            });
        }

        protected virtual void UpdateChapterPage(XDocument chapterPageDoc)
        {            
            if (!GenerateXmlOnly)
                OneNoteApp.UpdatePageContent(chapterPageDoc.ToString(), DateTime.MinValue, Constants.CurrentOneNoteSchema);                                 
        }

        protected virtual XDocument AddPage(string bookSectionId, string pageTitle, int pageLevel, out XmlNamespaceManager xnm)
        {
            xnm = null;

            if (GenerateXmlOnly)           
                return null;            

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
            if (GenerateXmlOnly)
                return null;

            return NotebookGenerator.AddTableToPage(chapterDoc, false, xnm, new CellInfo(SettingsManager.Instance.PageWidth_Bible), new CellInfo(NotebookGenerator.MinimalCellWidth));
        }

        protected virtual void AddVerseRowToTable(XElement tableElement, int verseNumber, int? topVerseNumber, string verseText)
        {
            if (!GenerateXmlOnly)
                NotebookGenerator.AddVerseRowToTable(tableElement, BibleBookContent.GetFullVerseString(verseNumber, topVerseNumber, verseText), 1, Locale);

            try
            {
                AddNewVerseContent(verseNumber, topVerseNumber, verseText);
            }
            catch (ConverterExceptionBase ex)
            {
                Errors.Add(ex);
            }
        }

        private void AddNewVerseContent(int verseNumber, int? topVerseNumber, string verseText)
        {
            var currentBook = BibleInfo.Content.Books.Last();
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
                    currentChapter.Verses.Add(new BibleVerseContent()
                    {
                        Index = i,
                        Value = string.Empty
                    });
                }

                throwException = true;                
            }

            currentChapter.Verses.Add(new BibleVerseContent()
            {
                Index = verseNumber,
                TopIndex = topVerseNumber,
                Value = verseText                
            });

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

            var module = new ModuleInfo() 
            { 
                ShortName = ModuleShortName,
                Name = extModuleInfo.Name, 
                Version = this.Version, 
                Notebooks = NotebooksInfo, 
                Type = IsStrong ? ModuleType.Strong: ModuleType.Bible 
            };
            module.BibleTranslationDifferences = this.TranslationDifferences;
            module.BibleStructure = new BibleStructureInfo()
            {
                Alphabet = extModuleInfo.Alphabet,                
                ChapterSectionNameTemplate = ChapterSectionNameTemplate              
            };
            module.Sections = this.SectionsInfo;
            module.DictionarySectionGroupName = this.DictionarySectionGroupName;
            module.StrongNumbersCount = this.StrongNumbersCount;

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
