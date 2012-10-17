using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using BibleCommon.Helpers;
using Microsoft.Office.Interop.OneNote;
using BibleCommon.Contracts;
using BibleCommon.Common;
using System.Xml.XPath;
using System.Xml;
using System.Xml.Linq;

namespace BibleCommon.Services
{
    /// <summary>
    /// На данный момент кэш не используется по следующим причинам:
    /// 1. Самое важное - долго дессириализуется! 4,5 секунд на хорошем компе!!!!! Надо оптимизировать структуру хранения данных. Возможно надо хранить дерево объектов, чтобы не дублировать id страниц и секций. + чтобы не хранить значения енумов. 
    /// 2. Надо доделать
    ///     - сейчас не хранятся ссылки на элементы. 
    ///     - соответственно везде, где после HierarchySearchManager.GetHierarchyObject() вызывается OneNoteUtils.GenerateHref() нужно доставать данные из кэша.
    /// </summary>
    public static class BibleVersesLinksCacheManager
    {
        private static string GetCacheFilePath(string notebookId)
        {            
            return Path.Combine(Utils.GetCacheFolderPath(), notebookId) + "_verses.cache";
        }

        public static bool CacheIsActive(string notebookId)
        {
            return File.Exists(GetCacheFilePath(notebookId));
        }

        public static VersePointersCachedLinks LoadBibleVersesLinks(string notebookId)
        {
            string filePath = GetCacheFilePath(notebookId);
            if (!File.Exists(filePath))
                throw new NotConfiguredException(string.Format("The file with Bible verses links does not exist: '{0}'", filePath));
            
            return (VersePointersCachedLinks)BinarySerializerHelper.Deserialize(filePath);
        }

        public static void GenerateBibleVersesLinks(Application oneNoteApp, string notebookId, string sectionGroupId, ICustomLogger logger)
        {
            string filePath = GetCacheFilePath(notebookId);
            if (File.Exists(filePath))
                throw new InvalidOperationException(string.Format("The file with Bible verses links already exists: '{0}'", filePath));

            var xnm = OneNoteUtils.GetOneNoteXNM();
            var result = new VersePointersCachedLinks();

            using (NotebookIterator iterator = new NotebookIterator(oneNoteApp))
            {
                BibleCommon.Services.NotebookIterator.NotebookInfo notebook = iterator.GetNotebookPages(notebookId, sectionGroupId, null);

                IterateContainer(oneNoteApp, notebook.RootSectionGroup, ref result, xnm, logger);
            }

            BinarySerializerHelper.Serialize(result, filePath);
        }

        private static void IterateContainer(Application oneNoteApp, NotebookIterator.SectionGroupInfo sectionGroup,
            ref VersePointersCachedLinks result, XmlNamespaceManager xnm, ICustomLogger logger)
        {
            foreach (NotebookIterator.SectionInfo section in sectionGroup.Sections)
            {
                BibleCommon.Services.Logger.LogMessage("section: " + section.Title);

                foreach (NotebookIterator.PageInfo page in section.Pages)
                {
                    logger.LogMessage(page.Title);

                    BibleCommon.Services.Logger.LogMessage("page: " + page.Title);

                    ProcessPage(oneNoteApp, page, section, ref result);
                }
            }

            foreach (NotebookIterator.SectionGroupInfo subSectionGroup in sectionGroup.SectionGroups)
            {
                IterateContainer(oneNoteApp, subSectionGroup, ref result, xnm, logger);
            }
        }

        private static void ProcessPage(Application oneNoteApp, NotebookIterator.PageInfo page, NotebookIterator.SectionInfo section,
            ref VersePointersCachedLinks result)
        {
            int? chapterNumber = StringUtils.GetStringFirstNumber(page.Title);
            if (!chapterNumber.HasValue)
                return;

            XmlNamespaceManager xnm;
            var pageId = (string)page.PageElement.Attribute("ID");
            var pageDoc = OneNoteUtils.GetPageContent(oneNoteApp, pageId, out xnm);

            var tableEl = NotebookGenerator.GetPageTable(pageDoc, xnm);
            if (tableEl == null)
                return;

            AddChapterPointer(oneNoteApp, section, pageDoc, pageId, chapterNumber, ref result, xnm);          
            
            foreach (var cellTextEl in tableEl.XPathSelectElements("one:Row/one:Cell[1]/one:OEChildren/one:OE/one:T", xnm))
            {
                AddVersePointer(oneNoteApp, section, pageDoc, pageId, chapterNumber, cellTextEl, ref result, xnm);          
            }
        }

        private static void AddItemToResult(Application oneNoteApp, VersePointer versePointer, NotebookIterator.SectionInfo section,
            string pageId, XElement objectEl, bool isChapter, ref VersePointersCachedLinks result)
        {
            if (versePointer.IsValid)
            {
                var commonKey = versePointer.ToSimpleVersePointer();

                foreach(var key in commonKey.GetAllVerses())
                {
                    if (!result.ContainsKey(key))
                    {
                        string textElId = (string)objectEl.Parent.Attribute("objectID");
                        string verseLink = SettingsManager.Instance.UseMiddleLinks ? null : OneNoteProxy.Instance.GenerateHref(oneNoteApp, pageId, textElId);

                        result.Add(key, new VersePointerLink()
                        {
                            SectionId = section.Id,
                            PageId = pageId,
                            ObjectId = textElId,
                            Href = verseLink,
                            VerseNumber = commonKey.VerseNumber,
                            IsChapter = isChapter
                        });
                    }
                }                
            }       
        }

        private static void AddVersePointer(Application oneNoteApp, NotebookIterator.SectionInfo section, 
            XDocument pageDoc, string pageId, int? chapterNumber, XElement cellTextEl,
            ref VersePointersCachedLinks result, XmlNamespaceManager xnm)
        {
            var verseNumber = VerseNumber.GetFromVerseText(cellTextEl.Value);                

            if (verseNumber.HasValue)
            {
                VersePointer versePointer = new VersePointer(section.Title, chapterNumber.Value, verseNumber.Value.Verse, verseNumber.Value.TopVerse);
                if (!versePointer.IsValid)
                    versePointer = new VersePointer(section.Title.Substring(4), chapterNumber.Value, verseNumber.Value.Verse, verseNumber.Value.TopVerse);  // иначе не понимает такие строки как "09. 1-я Царств 1:1"

                AddItemToResult(oneNoteApp, versePointer, section, pageId, cellTextEl, false, ref result);
            }      
        }

        private static void AddChapterPointer(Application oneNoteApp, NotebookIterator.SectionInfo section, 
            XDocument pageDoc, string pageId, int? chapterNumber,
            ref VersePointersCachedLinks result, XmlNamespaceManager xnm)
        {
            var pageTitleEl = NotebookGenerator.GetPageTitle(pageDoc, xnm);
            if (pageTitleEl != null)
            {
                VersePointer chapterPointer = new VersePointer(section.Title, chapterNumber.Value);
                if (!chapterPointer.IsValid)
                    chapterPointer = new VersePointer(section.Title.Substring(4), chapterNumber.Value);

                AddItemToResult(oneNoteApp, chapterPointer, section, pageId, pageTitleEl, true, ref result);                
            }
        }
    }
}
