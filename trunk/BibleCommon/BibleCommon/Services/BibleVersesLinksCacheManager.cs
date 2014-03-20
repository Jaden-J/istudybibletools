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
            return Path.Combine(Utils.GetCacheFolderPath(), notebookId) + "_verses" + Consts.Constants.FileExtensionCache;
        }

        public static bool CacheIsActive(string notebookId)
        {
            if (string.IsNullOrEmpty(notebookId))
                return false;

            return File.Exists(GetCacheFilePath(notebookId));
        }

        public static Dictionary<string, string> LoadBibleVersesLinks(string notebookId)
        {
            string filePath = GetCacheFilePath(notebookId);
            if (!File.Exists(filePath))
                throw new NotConfiguredException(string.Format("The file with Bible verses links does not exist: '{0}'", filePath));

            return SharpSerializationHelper.Deserialize<Dictionary<string, string>>(filePath);
        }

        public static void RemoveCacheFile(string notebookId)
        {
            if (!string.IsNullOrEmpty(notebookId))
            {
                var filePath = GetCacheFilePath(notebookId);
                File.Delete(filePath);
            }
        }

        public static void GenerateBibleVersesLinks(ref Application oneNoteApp, string notebookId, string sectionGroupId, bool toGenerateHref, ICustomLogger logger)
        {
            string filePath = GetCacheFilePath(notebookId);
            if (File.Exists(filePath))
                throw new InvalidOperationException(string.Format("The file with Bible verses links already exists: '{0}'", filePath));

            var xnm = OneNoteUtils.GetOneNoteXNM();
            var result = new Dictionary<string, string>();

            var iterator = new NotebookIterator();            
            BibleCommon.Services.NotebookIterator.NotebookInfo notebook = iterator.GetSectionGroupOrNotebookPages(ref oneNoteApp, notebookId, sectionGroupId, null);
            IterateContainer(ref oneNoteApp, notebookId, toGenerateHref, notebook.RootSectionGroup, ref result, xnm, logger);           

            SharpSerializationHelper.Serialize(result, filePath);
        }

        private static void IterateContainer(ref Application oneNoteApp, string notebookId, bool toGenerateHref, NotebookIterator.SectionGroupInfo sectionGroup,
            ref Dictionary<string, string> result, XmlNamespaceManager xnm, ICustomLogger logger)
        {
            foreach (NotebookIterator.SectionInfo section in sectionGroup.Sections)
            {
                BibleCommon.Services.Logger.LogMessage("section: " + section.Title);
                
                foreach (NotebookIterator.PageInfo page in section.Pages)
                {
                    logger.LogMessage(page.Title);

                    BibleCommon.Services.Logger.LogMessage("page: " + page.Title);

                    ProcessPage(ref oneNoteApp, notebookId, toGenerateHref, page, section, ref result);
                }
            }
            
            foreach (NotebookIterator.SectionGroupInfo subSectionGroup in sectionGroup.SectionGroups)
            {
                IterateContainer(ref oneNoteApp, notebookId, toGenerateHref, subSectionGroup, ref result, xnm, logger);
            }
        }

        private static void ProcessPage(ref Application oneNoteApp, string notebookId, bool toGenerateHref, NotebookIterator.PageInfo page, NotebookIterator.SectionInfo section, 
            ref Dictionary<string, string> result)
        {
            if (!section.Title.Contains(page.Title))   // иначе эта страница книги
            {
                int? chapterNumber = StringUtils.GetStringFirstNumber(page.Title);
                if (!chapterNumber.HasValue)
                    return;

                XmlNamespaceManager xnm;
                var pageId = (string)page.PageElement.Attribute("ID");
                var pageName = (string)page.PageElement.Attribute("name");
                var pageDoc = OneNoteUtils.GetPageContent(ref oneNoteApp, pageId, out xnm);

                AddChapterPointer(ref oneNoteApp, notebookId, toGenerateHref, section, pageDoc, pageId, pageName, chapterNumber, ref result, xnm);

                foreach (var cellTextEl in pageDoc.Root.XPathSelectElements("//one:Table/one:Row/one:Cell[1]/one:OEChildren/one:OE/one:T", xnm))
                {
                    AddVersePointer(ref oneNoteApp, notebookId, toGenerateHref, section, pageDoc, pageId, pageName, chapterNumber, cellTextEl, ref result, xnm);
                }
            }
        }

        private static void AddItemToResult(ref Application oneNoteApp, VersePointer versePointer, string notebookId, bool toGenerateHref, NotebookIterator.SectionInfo section,
            string pageId, string pageName, XElement objectEl, bool isChapter, ref Dictionary<string, string> result)
        {
            if (versePointer.IsValid)
            {
                var commonKey = versePointer;

                foreach (var key in commonKey.GetAllVerses(ref oneNoteApp, new GetAllIncludedVersesArgs() { BibleNotebookId = notebookId, Force = true }).Verses)
                {
                    var keyString = key.ToFirstVerseString();
                    if (!result.ContainsKey(keyString))
                    {
                        string textElId = (string)objectEl.Parent.Attribute("objectID");
                        string verseLink = toGenerateHref ? ApplicationCache.Instance.GenerateHref(ref oneNoteApp, pageId, textElId, new LinkProxyInfo(false, false)) : null;

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

        private static void AddVersePointer(ref Application oneNoteApp, string notebookId, bool toGenerateHref, NotebookIterator.SectionInfo section,
            XDocument pageDoc, string pageId, string pageName, int? chapterNumber, XElement cellTextEl,
            ref Dictionary<string, string> result, XmlNamespaceManager xnm)
        {
            var verseNumber = VerseNumber.GetFromVerseText(cellTextEl.Value);                

            if (verseNumber.HasValue)
            {
                VersePointer versePointer = new VersePointer(GetBookNameFromSectionTitle(section.Title), chapterNumber.Value, verseNumber.Value.Verse, verseNumber.Value.TopVerse);                

                AddItemToResult(ref oneNoteApp, versePointer, notebookId, toGenerateHref, section, pageId, pageName, cellTextEl, false, ref result);
            }      
        }

        private static void AddChapterPointer(ref Application oneNoteApp, string notebookId, bool toGenerateHref, NotebookIterator.SectionInfo section,
            XDocument pageDoc, string pageId, string pageName, int? chapterNumber,
            ref Dictionary<string, string> result, XmlNamespaceManager xnm)
        {
            var pageTitleEl = NotebookGenerator.GetPageTitle(pageDoc, xnm);
            if (pageTitleEl != null)
            {
                VersePointer chapterPointer = new VersePointer(GetBookNameFromSectionTitle(section.Title), chapterNumber.Value);                

                AddItemToResult(ref oneNoteApp, chapterPointer, notebookId, toGenerateHref, section, pageId, pageName, pageTitleEl, true, ref result);                
            }
        }

        private static string GetBookNameFromSectionTitle(string sectionTitle)
        {
            if (sectionTitle.Length > 4)
            {
                if (sectionTitle[2] == '.' && char.IsDigit(sectionTitle[0]) && char.IsDigit(sectionTitle[1]))                   // иначе не понимает такие строки как "09. 1-я Царств 1:1"
                    sectionTitle = sectionTitle.Substring(4);
                //else if (sectionTitle[3] == '.' && char.IsDigit(sectionTitle[0]) && char.IsDigit(sectionTitle[1]) && char.IsDigit(sectionTitle[2]))   // это если книг будет больше 100
                //    sectionTitle = sectionTitle.Substring(5);
            }

            return sectionTitle;
        }
    }
}
