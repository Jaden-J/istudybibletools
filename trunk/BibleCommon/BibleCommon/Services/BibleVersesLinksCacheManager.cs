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
            return Path.Combine(Utils.GetProgramDirectory(), notebookId) + ".cache";
        }

        public static bool CacheIsActive(string notebookId)
        {
            return File.Exists(GetCacheFilePath(notebookId));
        }

        public static Dictionary<VersePointer, HierarchySearchManager.HierarchySearchResult> LoadBibleVersesLinks(string notebookId)
        {
            string filePath = GetCacheFilePath(notebookId);
            if (!File.Exists(filePath))
                throw new NotConfiguredException(string.Format("The file with Bible verses links does not exist: '{0}'", filePath));

            return (Dictionary<VersePointer, HierarchySearchManager.HierarchySearchResult>)BinarySerializerHelper.Deserialize(filePath);
        }

        public static void GenerateBibleVersesLinks(Application oneNoteApp, string notebookId, string sectionGroupId, ICustomLogger logger)
        {
            string filePath = GetCacheFilePath(notebookId);
            if (File.Exists(filePath))
                throw new InvalidOperationException(string.Format("The file with Bible verses links already exists: '{0}'", filePath));

            var xnm = OneNoteUtils.GetOneNoteXNM();
            var result = new Dictionary<VersePointer, HierarchySearchManager.HierarchySearchResult>();

            using (NotebookIterator iterator = new NotebookIterator(oneNoteApp))
            {
                BibleCommon.Services.NotebookIterator.NotebookInfo notebook = iterator.GetNotebookPages(notebookId, sectionGroupId, null);

                IterateContainer(oneNoteApp, notebook.RootSectionGroup, result, xnm, logger);
            }

            BinarySerializerHelper.Serialize(result, filePath);
        }

        private static void IterateContainer(Application oneNoteApp, NotebookIterator.SectionGroupInfo sectionGroup,
            Dictionary<VersePointer, HierarchySearchManager.HierarchySearchResult> result, XmlNamespaceManager xnm, ICustomLogger logger)
        {
            foreach (NotebookIterator.SectionInfo section in sectionGroup.Sections)
            {
                BibleCommon.Services.Logger.LogMessage("section: " + section.Title);

                foreach (NotebookIterator.PageInfo page in section.Pages)
                {
                    logger.LogMessage(page.Title);

                    BibleCommon.Services.Logger.LogMessage("page: " + page.Title);

                    ProcessPage(oneNoteApp, page, section, result);
                }
            }

            foreach (NotebookIterator.SectionGroupInfo subSectionGroup in sectionGroup.SectionGroups)
            {
                IterateContainer(oneNoteApp, subSectionGroup, result, xnm, logger);
            }
        }

        private static void ProcessPage(Application oneNoteApp, NotebookIterator.PageInfo page, NotebookIterator.SectionInfo section, 
            Dictionary<VersePointer, HierarchySearchManager.HierarchySearchResult> result)
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

            AddChapterPointer(oneNoteApp, pageDoc, pageId, chapterNumber, section, result, xnm);          

            int temp;
            foreach (var cellTextEl in tableEl.XPathSelectElements("one:Row/one:Cell[1]/one:OEChildren/one:OE/one:T", xnm))
            {
                string verseNumber = StringUtils.GetNextString(cellTextEl.Value, -1, new SearchMissInfo(0, SearchMissInfo.MissMode.CancelOnMissFound),
                    out temp, out temp, StringSearchIgnorance.None, StringSearchMode.SearchNumber);

                if (!string.IsNullOrEmpty(verseNumber))
                {
                    VersePointer versePointer = new VersePointer(section.Title, chapterNumber.Value, int.Parse(verseNumber));
                    if (!versePointer.IsValid)
                        versePointer = new VersePointer(section.Title.Substring(4), chapterNumber.Value, int.Parse(verseNumber));  // иначе не понимает такие строки как "09. 1-я Царств 1:1"

                    if (versePointer.IsValid)
                    {
                        if (!result.ContainsKey(versePointer))
                        {
                            string textElId = (string)cellTextEl.Parent.Attribute("objectID");
                            var verseLink = OneNoteProxy.Instance.GenerateHref(oneNoteApp, pageId, textElId);

                            result.Add(versePointer, new HierarchySearchManager.HierarchySearchResult()
                            {
                                ResultType = HierarchySearchManager.HierarchySearchResultType.Successfully,
                                HierarchyStage = HierarchySearchManager.HierarchyStage.ContentPlaceholder,
                                HierarchyObjectInfo = new HierarchySearchManager.HierarchyObjectInfo()
                                    {
                                        ContentObjectId = new HierarchySearchManager.VerseObjectInfo() { ContentObjectId = textElId },
                                        PageId = pageId,
                                        SectionId = section.Id,
                                    }
                            });
                        }                       
                    }                  
                }               
            }
        }

        private static void AddChapterPointer(Application oneNoteApp, XDocument pageDoc, string pageId, int? chapterNumber, 
            NotebookIterator.SectionInfo section, Dictionary<VersePointer, HierarchySearchManager.HierarchySearchResult> result, XmlNamespaceManager xnm)
        {
            var pageTitleEl = NotebookGenerator.GetPageTitle(pageDoc, xnm);
            if (pageTitleEl != null)
            {
                VersePointer chapterPointer = new VersePointer(section.Title, chapterNumber.Value);
                if (!chapterPointer.IsValid)
                    chapterPointer = new VersePointer(section.Title.Substring(4), chapterNumber.Value);

                if (chapterPointer.IsValid)
                {
                    if (!result.ContainsKey(chapterPointer))
                    {
                        var pageTitleId = (string)pageTitleEl.Parent.Attribute("objectID");
                        var chapterLink = OneNoteProxy.Instance.GenerateHref(oneNoteApp, pageId, pageTitleId);

                        result.Add(chapterPointer, new HierarchySearchManager.HierarchySearchResult()
                        {
                            ResultType = HierarchySearchManager.HierarchySearchResultType.Successfully,
                            HierarchyStage = HierarchySearchManager.HierarchyStage.Page,
                            HierarchyObjectInfo = new HierarchySearchManager.HierarchyObjectInfo()
                            {
                                ContentObjectId = pageTitleId,
                                PageId = pageId,
                                SectionId = section.Id,
                            }
                        });
                    }                  
                }
            }
        }
    }
}
