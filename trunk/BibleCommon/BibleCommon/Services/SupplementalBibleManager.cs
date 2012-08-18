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

namespace BibleCommon.Services
{
    public static class SupplementalBibleManager
    {
        public static void CreateSupplementalBible(Application oneNoteApp, string moduleShortName)
        {
            if (!string.IsNullOrEmpty(SettingsManager.Instance.NotebookId_SupplementalBible))
            {
                if (!OneNoteUtils.NotebookExists(oneNoteApp, SettingsManager.Instance.NotebookId_SupplementalBible))
                    SettingsManager.Instance.NotebookId_SupplementalBible = null;
            }

            if (string.IsNullOrEmpty(SettingsManager.Instance.NotebookId_SupplementalBible))
            {
                SettingsManager.Instance.NotebookId_SupplementalBible = NotebookGenerator.CreateNotebook(oneNoteApp, Resources.Constants.SupplementalBibleName);
                SettingsManager.Instance.Save();
            }

            XDocument currentChapterDoc = null;
            XElement currentTableElement = null;
            string currentSectionGroupId = null;

            var moduleInfo = ModulesManager.GetModuleInfo(moduleShortName);
            var bibleInfo = ModulesManager.GetModuleBibleInfo(moduleShortName);

            for (int i = 0; i < moduleInfo.BibleStructure.BibleBooks.Count; i++)
            {
                var bibleBookInfo = moduleInfo.BibleStructure.BibleBooks[i];
                bibleBookInfo.SectionName = NotebookGenerator.GetBibleBookSectionName(bibleBookInfo.Name, i, moduleInfo.BibleStructure.OldTestamentBooksCount);

                if (string.IsNullOrEmpty(currentSectionGroupId))
                    currentSectionGroupId 
                        = NotebookGenerator.AddRootSectionGroupToNotebook(oneNoteApp, 
                            SettingsManager.Instance.NotebookId_SupplementalBible, moduleInfo.BibleStructure.OldTestamentName).Attribute("ID").Value;
                else if (i == moduleInfo.BibleStructure.OldTestamentBooksCount)
                    currentSectionGroupId 
                        = NotebookGenerator.AddRootSectionGroupToNotebook(oneNoteApp, 
                            SettingsManager.Instance.NotebookId_SupplementalBible, moduleInfo.BibleStructure.NewTestamentName).Attribute("ID").Value;

                var bookSectionId = NotebookGenerator.AddBookSectionToBibleNotebook(oneNoteApp, currentSectionGroupId, bibleBookInfo.SectionName, bibleBookInfo.Name);

                var bibleBook = bibleInfo.Content.Books.FirstOrDefault(book => book.Index == bibleBookInfo.Index);
                if (bibleBook == null)
                    throw new Exception("Manifest.xml has Bible books that does not exists in bible.xml");

                foreach (var chapter in bibleBook.Chapters)
                {
                    if (currentChapterDoc != null)
                        UpdateChapterPage(oneNoteApp, currentChapterDoc);

                    string chapterSectionName = string.Format(moduleInfo.BibleStructure.ChapterSectionNameTemplate, chapter.Index, bibleBookInfo.Name);

                    XmlNamespaceManager xnm;
                    currentChapterDoc = NotebookGenerator.AddChapterPageToBibleNotebook(oneNoteApp, bookSectionId, chapterSectionName, 1, bibleInfo.Content.Locale, out xnm);

                    currentTableElement = NotebookGenerator.AddTableToBibleChapterPage(currentChapterDoc, SettingsManager.Instance.PageWidth_Bible, xnm);

                    foreach (var verse in chapter.Verses)
                    {
                        NotebookGenerator.AddVerseRowToBibleTable(currentTableElement, string.Format("{0} {1}", verse.Index, verse.Value), bibleInfo.Content.Locale);            
                    }
                }                
            }

            if (currentChapterDoc != null)
            {
                UpdateChapterPage(oneNoteApp, currentChapterDoc);
            }
        }

        private static void UpdateChapterPage(Application oneNoteApp, XDocument chapterPageDoc)
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
