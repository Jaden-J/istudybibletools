using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using BibleCommon.Helpers;

namespace BibleConfigurator.ModuleConverter
{
    public class BibleQuotaModuleInfo : ExternalModuleInfo
    {
        public string ChapterSign { get; set; }
        public string VerseSign { get; set; }

        public List<BibleQuotaBibleBookInfo> BibleBooksInfo { get; set; } 

        public BibleQuotaModuleInfo()
        {
            BibleBooksInfo = new List<BibleQuotaBibleBookInfo>();
        }
    }

    public class BibleQuotaBibleBookInfo 
    {
        public string Name { get; set; }
        public List<string> Shortenings { get; set; }
        public int ChaptersCount { get; set; }
        public string FileName { get; set; }
    }

    public class BibleQuotaConverter: ConverterBase
    {
        protected const string IniFileName = "bibleqt.ini";

        protected string ModuleFolder { get; set; }

        public BibleQuotaConverter(string emptyNotebookName, string moduleFolder)
            : base(emptyNotebookName)
        {
            this.ModuleFolder = moduleFolder;
        }


        protected override ExternalModuleInfo ReadExternalModuleInfo()
        {
            string iniFilePath = Path.Combine(ModuleFolder, IniFileName);
            string fileContent = File.ReadAllText(iniFilePath);

            var result = new BibleQuotaModuleInfo();

            foreach(string line in fileContent.Split(new char[] { '\n' }))
            {
                var pair = line.Split(new char[] { '=' }, 2, StringSplitOptions.RemoveEmptyEntries);
                if (pair.Length == 2)
                {
                    string key = pair[0].Trim();
                    string value = pair[1].Trim();

                    if (key == "BibleName")
                        result.Name = value;
                    else if (key == "BibleShortName")
                        result.ShortName = value;
                    else if (key == "Alphabet")
                        result.Alphabet = value;
                    else if (key == "BookQty")
                        result.BooksCount = int.Parse(value);
                    else if (key == "ChapterSign")
                        result.ChapterSign = value;
                    else if (key == "VerseSign")
                        result.VerseSign = value;
                    else if (key == "PathName")
                        result.BibleBooksInfo.Add(new BibleQuotaBibleBookInfo() { FileName = value });
                    else if (key == "FullName")
                        result.BibleBooksInfo[result.BibleBooksInfo.Count - 1].Name = value;
                    else if (key == "ShortName")
                        result.BibleBooksInfo[result.BibleBooksInfo.Count - 1].Shortenings = value.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                    else if (key == "ChapterQty")
                        result.BibleBooksInfo[result.BibleBooksInfo.Count - 1].ChaptersCount = int.Parse(value);
                }
            }

            return result;
        }

        protected override void ProcessBibleBooks(ExternalModuleInfo externalModuleInfo)
        {
            var moduleInfo = (BibleQuotaModuleInfo)externalModuleInfo;

            foreach (var bibleBookInfo in moduleInfo.BibleBooksInfo)
            {
                var bookSectionId = AddNewBook(bibleBookInfo.Name);

                string bookFile = Path.Combine(ModuleFolder, bibleBookInfo.FileName);
                string fileContent = File.ReadAllText(bookFile);

                foreach (string line in fileContent.Split(new char[] { '\n' }))
                {
                    if (line.StartsWith(moduleInfo.ChapterSign))
                    {
                       string chapterTitle = StringUtils.GetText(line, moduleInfo.Alphabet).Trim();

                       var chpaterDoc = AddNewChapter(bookSectionId, chapterTitle);
                    }
                }
            }
        }
    }
}
