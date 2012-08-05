using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using BibleCommon.Common;
using Microsoft.Office.Interop.OneNote;
using BibleCommon.Helpers;
using System.Xml;

namespace BibleCommon.Services
{
    public static class BibleParallelTranslationManager
    {
        private const string supportedModuleMinVersion = "2.0";

        private enum BibleGeneratorDecision
        {
            SameLocation,

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="moduleShortName">Module Directory Name</param>
        /// <param name="translationIndex">0 - основная Библия</param>
        public static void AddParallelTranslation(Application oneNoteApp, string moduleShortName)
        {
            if (!SettingsManager.Instance.CurrentModuleIsCorrect())
                throw new InvalidOperationException("Current module is not correct.");

            if (SettingsManager.Instance.CurrentModule.Version.CompareTo(supportedModuleMinVersion) < 0)
                throw new NotSupportedException(string.Format("Version of current module is {0}.", SettingsManager.Instance.CurrentModule.Version));            
            
            var parallelModuleInfo = ModulesManager.GetModuleInfo(moduleShortName);
            ModulesManager.CheckModule(parallelModuleInfo);  // если с модулем то-то не так, то выдаст ошибку
            if (parallelModuleInfo.Version.CompareTo(supportedModuleMinVersion) < 0)
                throw new NotSupportedException(string.Format("Version of parallel module is {0}.", SettingsManager.Instance.CurrentModule.Version));


            var translationIndex = SettingsManager.Instance.ParallelModules.Count;
            var baseBibleInfo = ModulesManager.GetModuleBibleInfo(SettingsManager.Instance.ModuleName);
            var parallelBibleInfo = ModulesManager.GetModuleBibleInfo(moduleShortName);            

            GenerateParallelBibleTables(oneNoteApp, SettingsManager.Instance.NotebookId_Bible, SettingsManager.Instance.CurrentModule, baseBibleInfo, parallelModuleInfo, parallelBibleInfo);            
        }

        private static void GenerateParallelBibleTables(Application oneNoteApp, string bibleNotebookId, ModuleInfo baseModuleInfo, ModuleBibleInfo baseBibleInfo, 
            ModuleInfo parallelModuleInfo, ModuleBibleInfo parallelBibleInfo)
        {
            foreach (var baseBibleBook in baseBibleInfo.Content.Books)
            {
                var baseBookInfo = baseModuleInfo.BibleStructure.BibleBooks.FirstOrDefault(b => b.Index == baseBibleBook.Index);
                if (baseBookInfo == null)
                    throw new InvalidModuleException(string.Format("Book with index {0} is not found in module manifest", baseBibleBook.Index));

                var parallelBibleBook = parallelBibleInfo.Content.Books.FirstOrDefault(b => b.Index == baseBibleBook.Index);
                if (parallelBibleBook != null)
                {
                    XElement sectionEl = HierarchySearchManager.FindBibleBookSection(oneNoteApp, bibleNotebookId, baseBookInfo.SectionName);
                    if (sectionEl == null)
                        throw new Exception(string.Format("Section with name {0} is not found", baseBookInfo.SectionName));

                    var test = new BibleTranslationDifferencesEx(parallelBibleInfo.TranslationDifferences);

                    var baseTranslationDifferences = baseBibleInfo.TranslationDifferences.BookDifferences.FirstOrDefault(b => b.BookIndex == baseBibleBook.Index);
                    var parallelTranslationDifferences = parallelBibleInfo.TranslationDifferences.BookDifferences.FirstOrDefault(b => b.BookIndex == baseBibleBook.Index);

                    ProcessBibleBook(oneNoteApp, sectionEl, baseBookInfo, baseBibleBook, baseTranslationDifferences, parallelBibleBook, parallelTranslationDifferences);                    
                }
            }
        }

        private static void ProcessBibleBook(Application oneNoteApp, XElement sectionEl, BibleBookInfo baseBibleInfo, BibleBookContent baseBibleBook, BibleBookDifferences baseTranslationDifferences,
            BibleBookContent parallelBibleBook, BibleBookDifferences parallelTranslationDifferences)
        {
            XmlNamespaceManager xnm;
            string sectionId = (string)sectionEl.Attribute("ID");
            string sectionName = (string)sectionEl.Attribute("name");

            var sectionPages = OneNoteUtils.GetHierarchyElement(oneNoteApp, sectionId, HierarchyScope.hsPages, out xnm);            

            foreach (var baseChapter in baseBibleBook.Chapters)
            {
                
            }

            foreach (var pageEl in sectionPages.Root.Elements())
            {
                string pageId = (string)pageEl.Attribute("ID");
                string pageName = (string)pageEl.Attribute("name");

                int? chapter = GetChapter(pageName, baseBibleInfo.Name);                

                if (chapter.HasValue)
                {
                    var chapterDoc = OneNoteUtils.GetPageContent(oneNoteApp, pageId, out xnm);

                    NotebookGenerator.AddTableToBibleChapterPage(chapterDoc, SettingsManager.Instance.PageWidth_Bible, xnm);

                    
                }
            }
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

        public static void RemoveLastParallelTranslation(Application oneNoteApp)
        {
            throw new NotImplementedException();
        }
    }
}
