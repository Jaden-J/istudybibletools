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
using Polenter.Serialization;

namespace BibleCommon.Services
{   
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

        public static Dictionary<string, string> LoadBibleVersesLinks(string notebookId)
        {
            string filePath = GetCacheFilePath(notebookId);
            if (!File.Exists(filePath))
                throw new NotConfiguredException(string.Format("The file with Bible verses links does not exist: '{0}'", filePath));

            return SharpSerializationHelper.Deserialize<Dictionary<string, string>>(filePath);
        }

        public static void GenerateBibleVersesLinks(ref Application oneNoteApp, string notebookId, string sectionGroupId, ICustomLogger logger)
        {
            string filePath = GetCacheFilePath(notebookId);
            if (File.Exists(filePath))
                throw new InvalidOperationException(string.Format("The file with Bible verses links already exists: '{0}'", filePath));

            var xnm = OneNoteUtils.GetOneNoteXNM();
            var result = new Dictionary<string, string>();

            using (NotebookIterator iterator = new NotebookIterator(oneNoteApp))
            {
                BibleCommon.Services.NotebookIterator.NotebookInfo notebook = iterator.GetNotebookPages(notebookId, sectionGroupId, null);

                IterateContainer(ref oneNoteApp, notebookId, notebook.RootSectionGroup, ref result, xnm, logger);
            }

            SharpSerializationHelper.Serialize(result, filePath);
        }

        private static void IterateContainer(ref Application oneNoteApp, string notebookId, NotebookIterator.SectionGroupInfo sectionGroup,
            ref Dictionary<string, string> result, XmlNamespaceManager xnm, ICustomLogger logger)
        {
            foreach (NotebookIterator.SectionInfo section in sectionGroup.Sections)
            {
                BibleCommon.Services.Logger.LogMessageParams("section: " + section.Title);

                foreach (NotebookIterator.PageInfo page in section.Pages)
                {
                    logger.LogMessage(page.Title);

                    BibleCommon.Services.Logger.LogMessageParams("page: " + page.Title);

                    ProcessPage(ref oneNoteApp, notebookId, page, section, ref result);
                }
            }
            
            foreach (NotebookIterator.SectionGroupInfo subSectionGroup in sectionGroup.SectionGroups)
            {
                IterateContainer(ref oneNoteApp, notebookId, subSectionGroup, ref result, xnm, logger);
            }
        }

        private static void ProcessPage(ref Application oneNoteApp, string notebookId, NotebookIterator.PageInfo page, NotebookIterator.SectionInfo section,
            ref Dictionary<string, string> result)
        {
            int? chapterNumber = StringUtils.GetStringFirstNumber(page.Title);
            if (!chapterNumber.HasValue)
                return;

            XmlNamespaceManager xnm;
            var pageId = (string)page.PageElement.Attribute("ID");
            var pageName = (string)page.PageElement.Attribute("name");
            var pageDoc = OneNoteUtils.GetPageContent(ref oneNoteApp, pageId, out xnm);

            var tableEl = NotebookGenerator.GetPageTable(pageDoc, xnm);
            if (tableEl == null)
                return;

            AddChapterPointer(ref oneNoteApp, notebookId, section, pageDoc, pageId, pageName, chapterNumber, ref result, xnm);          
            
            foreach (var cellTextEl in tableEl.XPathSelectElements("one:Row/one:Cell[1]/one:OEChildren/one:OE/one:T", xnm))
            {
                AddVersePointer(ref oneNoteApp, notebookId, section, pageDoc, pageId, pageName, chapterNumber, cellTextEl, ref result, xnm);          
            }
        }

        private static void AddItemToResult(ref Application oneNoteApp, VersePointer versePointer, string notebookId, NotebookIterator.SectionInfo section,
            string pageId, string pageName, XElement objectEl, bool isChapter, ref Dictionary<string, string> result)
        {
            if (versePointer.IsValid)
            {
                var commonKey = versePointer;

                foreach (var key in commonKey.GetAllVerses(ref oneNoteApp, new GetAllIncludedVersesExceptFirstArgs() { BibleNotebookId = notebookId, Force = true }))
                {
                    var keyString = key.ToFirstVerseString();
                    if (!result.ContainsKey(keyString))
                    {
                        string textElId = (string)objectEl.Parent.Attribute("objectID");
                        string verseLink = OneNoteProxy.Instance.GenerateHref(ref oneNoteApp, pageId, textElId);

                        result.Add(keyString, new VersePointerLink()
                        {
                            SectionId = section.Id,
                            PageId = pageId,
                            PageName = pageName,
                            ObjectId = textElId,
                            Href = verseLink,
                            VerseNumber = commonKey.VerseNumber,
                            IsChapter = isChapter
                        }.ToString());
                    }
                }                
            }       
        }

        private static void AddVersePointer(ref Application oneNoteApp, string notebookId, NotebookIterator.SectionInfo section,
            XDocument pageDoc, string pageId, string pageName, int? chapterNumber, XElement cellTextEl,
            ref Dictionary<string, string> result, XmlNamespaceManager xnm)
        {
            var verseNumber = VerseNumber.GetFromVerseText(cellTextEl.Value);                

            if (verseNumber.HasValue)
            {
                VersePointer versePointer = new VersePointer(section.Title, chapterNumber.Value, verseNumber.Value.Verse, verseNumber.Value.TopVerse);
                if (!versePointer.IsValid)
                    versePointer = new VersePointer(section.Title.Substring(4), chapterNumber.Value, verseNumber.Value.Verse, verseNumber.Value.TopVerse);  // иначе не понимает такие строки как "09. 1-я Царств 1:1"

                AddItemToResult(ref oneNoteApp, versePointer, notebookId, section, pageId, pageName, cellTextEl, false, ref result);
            }      
        }

        private static void AddChapterPointer(ref Application oneNoteApp, string notebookId, NotebookIterator.SectionInfo section,
            XDocument pageDoc, string pageId, string pageName, int? chapterNumber,
            ref Dictionary<string, string> result, XmlNamespaceManager xnm)
        {
            var pageTitleEl = NotebookGenerator.GetPageTitle(pageDoc, xnm);
            if (pageTitleEl != null)
            {
                VersePointer chapterPointer = new VersePointer(section.Title, chapterNumber.Value);
                if (!chapterPointer.IsValid)
                    chapterPointer = new VersePointer(section.Title.Substring(4), chapterNumber.Value);

                AddItemToResult(ref oneNoteApp, chapterPointer, notebookId, section, pageId, pageName, pageTitleEl, true, ref result);                
            }
        }
    }
}
