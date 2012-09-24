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

namespace BibleCommon.Services
{
    public static class BibleVersesLinksCacheManager
    {
        private static string GetCacheFilePath()
        {
            return Path.Combine(Utils.GetProgramDirectory(), SettingsManager.Instance.NotebookId_Bible + ".cache");
        }

        public static bool CacheIsActive()
        {
            return File.Exists(GetCacheFilePath());
        }

        public static Dictionary<SimpleVersePointer, string> LoadBibleVersesLinks()
        {
            string filePath = GetCacheFilePath();
            if (!File.Exists(filePath))
                throw new SystemNotConfiguredException( (string.Format("The file with Bible verses links does not exist: '{0}'", filePath));

            return (Dictionary<SimpleVersePointer, string>)BinarySerializerHelper.Deserialize(filePath);
        }     

        public static void GenerateBibleVersesLinks(Application oneNoteApp, ICustomLogger logger)
        {
            string filePath = GetCacheFilePath();
            if (File.Exists(filePath))
                throw new InvalidOperationException(string.Format("The file with Bible verses links already exists: '{0}'", filePath));

            var xnm = OneNoteUtils.GetOneNoteXNM();
            int temp;
            var result = new Dictionary<SimpleVersePointer, string>();
            var bibleIterator = new BibleParallelTranslationManager(oneNoteApp, SettingsManager.Instance.ModuleName, SettingsManager.Instance.ModuleName, SettingsManager.Instance.NotebookId_Bible);
            bibleIterator.Logger = logger;
            bibleIterator.IterateBaseBible(
                (chapterPageDoc, chapterPointer) =>
                {
                    var tableEl = NotebookGenerator.GetPageTable(chapterPageDoc, xnm);
                    var pageId = (string)chapterPageDoc.Root.Attribute("ID");

                    foreach (var cellTextEl in tableEl.XPathSelectElements("one:Row/one:Cell[1]/one:OEChildren/one:OE/one:T", xnm))
                    {
                        string verseNumber = StringUtils.GetNextString(cellTextEl.Value, -1, new SearchMissInfo(0, SearchMissInfo.MissMode.CancelOnMissFound),
                            out temp, out temp, StringSearchIgnorance.None, StringSearchMode.SearchNumber);

                        if (!string.IsNullOrEmpty(verseNumber))
                        {
                            var versePointer = new SimpleVersePointer(chapterPointer.BookIndex, chapterPointer.Chapter, int.Parse(verseNumber));

                            if (!result.ContainsKey(versePointer))
                            {
                                var verseLink = OneNoteProxy.Instance.GenerateHref(oneNoteApp, pageId, (string)cellTextEl.Attribute("objectID"));
                                result.Add(versePointer, verseLink);
                            }
                        }
                    }

                    return null;
                }, false, false, null);

            BinarySerializerHelper.Serialize(result, filePath);
        }

    }
}
