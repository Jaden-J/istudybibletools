using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.OneNote;
using System.Xml.Linq;
using BibleCommon.Helpers;
using System.IO;
using BibleCommon.Consts;
using System.Xml;
using BibleCommon.Common;
using System.Xml.XPath;

namespace BibleCommon.Services
{
    public static class SupplementalBibleManager
    {
        public static void CreateSupplementalBible(Application oneNoteApp, string moduleShortName)
        {
            if (!string.IsNullOrEmpty(SettingsManager.Instance.NotebookId_SupplementalBible))            
                if (!OneNoteUtils.NotebookExists(oneNoteApp, SettingsManager.Instance.NotebookId_SupplementalBible))
                    SettingsManager.Instance.NotebookId_SupplementalBible = null;
            

            if (string.IsNullOrEmpty(SettingsManager.Instance.NotebookId_SupplementalBible))            
                SettingsManager.Instance.NotebookId_SupplementalBible = NotebookGenerator.CreateNotebook(oneNoteApp, Resources.Constants.SupplementalBibleName);                            
            else
                throw new InvalidOperationException("Supplemental Bible already exists");

            SettingsManager.Instance.SupplementalBibleModules.Clear();
            SettingsManager.Instance.SupplementalBibleModules.Add(moduleShortName);
            SettingsManager.Instance.Save();
            
            string currentSectionGroupId = null;
            var moduleInfo = ModulesManager.GetModuleInfo(moduleShortName);
            var bibleInfo = ModulesManager.GetModuleBibleInfo(moduleShortName);            

            for (int i = 0; i < moduleInfo.BibleStructure.BibleBooks.Count; i++)
            {
                var bibleBookInfo = moduleInfo.BibleStructure.BibleBooks[i];
                bibleBookInfo.SectionName = NotebookGenerator.GetBibleBookSectionName(bibleBookInfo.Name, i, moduleInfo.BibleStructure.OldTestamentBooksCount);

                currentSectionGroupId = GetCurrentSectionGroupId(oneNoteApp, currentSectionGroupId, moduleInfo, i);                

                var bookSectionId = NotebookGenerator.AddBookSectionToBibleNotebook(oneNoteApp, currentSectionGroupId, bibleBookInfo.SectionName, bibleBookInfo.Name);

                var bibleBook = bibleInfo.Content.Books.FirstOrDefault(book => book.Index == bibleBookInfo.Index);
                if (bibleBook == null)
                    throw new Exception("Manifest.xml has Bible books that do not exist in bible.xml");

                foreach (var chapter in bibleBook.Chapters)
                {   
                    GenerateChapterPage(oneNoteApp, chapter, bookSectionId, moduleInfo, bibleBookInfo, bibleInfo);                    
                }                
            }            
        }

        public static void LinkSupplementalBibleWithMainBible(Application oneNoteApp, int supplementalModuleIndex)
        {



            var chapterPageId = (string)chapterPageDoc.Root.Attribute("ID");
            oneNoteApp.SyncHierarchy(chapterPageId);

            UpdateBibleChapterLinksToSupplementalBible(oneNoteApp, chapterPageId, chapterIndex, bibleBookInfo);

            OneNoteProxy.Instance.CommitAllModifiedPages(oneNoteApp, pageContent => pageContent.PageType == OneNoteProxy.PageType.Bible, null, null);
        }               

        private static void UpdateBibleChapterLinksToSupplementalBible(Application oneNoteApp, string chapterPageId, int chapterIndex, BibleBookInfo bibleBookInfo)
        {
            XmlNamespaceManager xnm;
            var nms = XNamespace.Get(Constants.OneNoteXmlNs);           
            
            var chapterPageDoc = OneNoteUtils.GetPageContent(oneNoteApp, chapterPageId, out xnm);
            var bibleTable = NotebookGenerator.GetBibleTable(chapterPageDoc, xnm);
            SimpleVersePointer prevParallelVersePointer = null;


            foreach(var cell in bibleTable.XPathSelectElements("one:Row/one:Cell[1]/one:OEChildren/one:OE/one:T", xnm).Skip(1))
            {
                var verseIndex = StringUtils.GetStringFirstNumber(cell.Value);
                if (verseIndex.HasValue)
                {
                    var cellId = (string)cell.Parent.Attribute("objectID").Value;
                    string link = OneNoteUtils.GenerateHref(oneNoteApp, SettingsManager.Instance.SupplementalBibleLinkName, chapterPageId, cellId);

                    var baseVersePointer = new SimpleVersePointer(bibleBookInfo.Index, chapterIndex, verseIndex.Value);
                    var parallelVersePointer = BibleParallelTranslationConnectorManager.GetParallelVersePointer(baseVersePointer,
                        SettingsManager.Instance.SupplementalBibleModules.First(), SettingsManager.Instance.ModuleName);

                    if (parallelVersePointer != prevParallelVersePointer)
                    {
                        prevParallelVersePointer = parallelVersePointer;
                        if (!parallelVersePointer.IsEmpty)
                        {
                            var mainBibleChapterDoc = PrepareMainBibleTable(oneNoteApp, parallelVersePointer, out xnm);

                            if (mainBibleChapterDoc != null)
                            {
                                var mainBibleVerseEl = HierarchySearchManager.FindVerse(mainBibleChapterDoc, false, parallelVersePointer.Verse, xnm);
                                if (mainBibleVerseEl == null)
                                    throw new Exception(string.Format("Can not find Bible verse cell for '{0} {1}:{2}'", bibleBookInfo.Name, chapterIndex, verseIndex));

                                var mainBibleVerseRowEl = mainBibleVerseEl.Parent.Parent.Parent.Parent;
                                var sbCell = mainBibleVerseRowEl.XPathSelectElement("one:Cell[3]/one:OEChildren/one:OE/one:T", xnm);

                                if (sbCell == null)
                                    mainBibleVerseRowEl.Add(NotebookGenerator.GetCell(link, string.Empty, nms));
                                else
                                    sbCell.Value = link;
                            }
                        }
                    }
                    else
                    {
                        int i = 0;
                    }
                }
            }            
        }

        private static XDocument PrepareMainBibleTable(Application oneNoteApp, SimpleVersePointer versePointer, out XmlNamespaceManager xnm)  // метод поддерживает кэширование
        {
            xnm = null;

            var mainBibleBookInfo = SettingsManager.Instance.CurrentModule.BibleStructure.BibleBooks.FirstOrDefault(book => book.Index == versePointer.BookIndex);
            if (mainBibleBookInfo != null)
            {

                var mainBibleChapterPageEl = HierarchySearchManager.FindChapterPage(oneNoteApp, SettingsManager.Instance.NotebookId_Bible, mainBibleBookInfo.SectionName, versePointer.Chapter);
                if (mainBibleChapterPageEl == null)
                    throw new Exception(string.Format("The Bible page for chapter {0} of book {1} does not found", versePointer.Chapter, mainBibleBookInfo.Name));
                string mainBibleChapterPageId = (string)mainBibleChapterPageEl.Attribute("ID");

                var mainBibleChapterPageDoc = OneNoteProxy.Instance.GetPageContent(oneNoteApp, mainBibleChapterPageId, OneNoteProxy.PageType.Bible);
                var mainBibleTableElement = NotebookGenerator.GetBibleTable(mainBibleChapterPageDoc.Content, mainBibleChapterPageDoc.Xnm);

                var columnsCount = mainBibleTableElement.XPathSelectElements("one:Columns/one:Column", mainBibleChapterPageDoc.Xnm).Count();
                if (columnsCount == 2)
                    NotebookGenerator.AddColumnToTable(mainBibleTableElement, NotebookGenerator.MinimalCellWidth, mainBibleChapterPageDoc.Xnm);

                xnm = mainBibleChapterPageDoc.Xnm;
                mainBibleChapterPageDoc.WasModified = true;

                return mainBibleChapterPageDoc.Content;
            }

            return null;
        }

        private static void GenerateChapterPage(Application oneNoteApp, BibleChapterContent chapter, string bookSectionId,
           ModuleInfo moduleInfo, BibleBookInfo bibleBookInfo, ModuleBibleInfo bibleInfo)
        {
            string chapterPageName = string.Format(moduleInfo.BibleStructure.ChapterSectionNameTemplate, chapter.Index, bibleBookInfo.Name);

            XmlNamespaceManager xnm;
            var currentChapterDoc = NotebookGenerator.AddChapterPageToBibleNotebook(oneNoteApp, bookSectionId, chapterPageName, 1, bibleInfo.Content.Locale, out xnm);

            var currentTableElement = NotebookGenerator.AddTableToBibleChapterPage(currentChapterDoc, SettingsManager.Instance.PageWidth_Bible, xnm);

            NotebookGenerator.AddParallelBibleTitle(currentTableElement, moduleInfo.Name, 0, bibleInfo.Content.Locale, xnm);

            foreach (var verse in chapter.Verses)
            {
                NotebookGenerator.AddVerseRowToBibleTable(currentTableElement, string.Format("{0} {1}", verse.Index, verse.Value), bibleInfo.Content.Locale);
            }

            UpdateChapterPage(oneNoteApp, currentChapterDoc, chapter.Index, bibleBookInfo);
        }       

        private static string GetCurrentSectionGroupId(Application oneNoteApp, string currentSectionGroupId, Common.ModuleInfo moduleInfo, int i)
        {
            if (string.IsNullOrEmpty(currentSectionGroupId))
                currentSectionGroupId
                    = NotebookGenerator.AddRootSectionGroupToNotebook(oneNoteApp,
                        SettingsManager.Instance.NotebookId_SupplementalBible, moduleInfo.BibleStructure.OldTestamentName).Attribute("ID").Value;
            else if (i == moduleInfo.BibleStructure.OldTestamentBooksCount)
                currentSectionGroupId
                    = NotebookGenerator.AddRootSectionGroupToNotebook(oneNoteApp,
                        SettingsManager.Instance.NotebookId_SupplementalBible, moduleInfo.BibleStructure.NewTestamentName).Attribute("ID").Value;

            return currentSectionGroupId;
        }

        private static void UpdateChapterPage(Application oneNoteApp, XDocument chapterPageDoc, int chapterIndex, BibleBookInfo bibleBookInfo)
        {
            oneNoteApp.UpdatePageContent(chapterPageDoc.ToString(), DateTime.MinValue, Constants.CurrentOneNoteSchema);            
        }

        public static BibleParallelTranslationConnectionResult AddParallelBible(Application oneNoteApp, string moduleShortName)
        {
            if (string.IsNullOrEmpty(SettingsManager.Instance.NotebookId_SupplementalBible))            
                throw new Exception(BibleCommon.Resources.Constants.Error_SystemIsNotConfigures);                            

            var result = BibleParallelTranslationManager.AddParallelTranslation(oneNoteApp, moduleShortName);            

            SettingsManager.Instance.SupplementalBibleModules.Add(moduleShortName);


            // ещё надо объединить сокращения книг

            SettingsManager.Instance.Save();

            return result;
        }
    }
}
