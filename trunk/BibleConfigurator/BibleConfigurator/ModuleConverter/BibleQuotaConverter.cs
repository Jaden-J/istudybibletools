using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using BibleCommon.Helpers;
using System.Xml.Linq;
using BibleCommon.Consts;
using System.Xml;
using System.Xml.XPath;
using System.Globalization;
using BibleCommon.Common;
using System.Xml.Serialization;

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
        public List<Abbreviation> Abbreviations { get; set; }
        public int ChaptersCount { get; set; }
        public string FileName { get; set; }
        public string SectionName { get; set; }
    }

    public class BibleQuotaConverter: ConverterBase
    {
        protected const string IniFileName = "bibleqt.ini";

        protected string ModuleFolder { get; set; }
        protected Encoding FileEncoding { get; set; }

        public Func<BibleQuotaBibleBookInfo, string, string> ConvertChapterNameFunc { get; set; }        
        

        /// <summary>
        /// 
        /// </summary>
        /// <param name="emptyNotebookName"></param>
        /// <param name="moduleFolder"></param>
        /// <param name="manifestFilePathToSave"></param>
        /// <param name="fileEncoding"></param>
        /// <param name="oldTestamentName"></param>
        /// <param name="newTestamentName"></param>
        /// <param name="oldTestamentBooksCount"></param>
        /// <param name="newTestamentBooksCount"></param>
        /// <param name="locale">can be not specified</param>
        /// <param name="notebooksInfo"></param>
        public BibleQuotaConverter(string emptyNotebookName, string moduleFolder, string manifestFilesFolderPath, Encoding fileEncoding,
            string oldTestamentName, string newTestamentName, int oldTestamentBooksCount, int newTestamentBooksCount,
            string locale, List<NotebookInfo> notebooksInfo, List<int> bookIndexes, BibleTranslationDifferences translationDifferences)
            : base(emptyNotebookName, manifestFilesFolderPath, oldTestamentName, newTestamentName, oldTestamentBooksCount, newTestamentBooksCount, 
                        locale, notebooksInfo, bookIndexes, translationDifferences)
        {
            this.ModuleFolder = moduleFolder;
            this.FileEncoding = fileEncoding;            
        }

        protected override ExternalModuleInfo ReadExternalModuleInfo()
        {
            string iniFilePath = Path.Combine(ModuleFolder, IniFileName);            

            var result = new BibleQuotaModuleInfo();

            foreach (string line in File.ReadAllLines(iniFilePath, FileEncoding))
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
                        result.BibleBooksInfo[result.BibleBooksInfo.Count - 1].Abbreviations 
                            = value
                                .ToLowerInvariant()
                                .Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)
                                .Select(s => s.Trim(new char[] { '.' })).Distinct()
                                .Select(s => new Abbreviation(s)).ToList();
                    else if (key == "ChapterQty")
                        result.BibleBooksInfo[result.BibleBooksInfo.Count - 1].ChaptersCount = int.Parse(value);
                }
            }

            return result;
        }

        protected override void ProcessBibleBooks(ExternalModuleInfo externalModuleInfo)
        {
            var moduleInfo = (BibleQuotaModuleInfo)externalModuleInfo;

            XDocument currentChapterDoc = null;
            XElement currentTableElement = null;            
            string currentSectionGroupId = null;

            for (int i = 0; i < moduleInfo.BibleBooksInfo.Count; i++) 
            {
                var bibleBookInfo = moduleInfo.BibleBooksInfo[i];                
                bibleBookInfo.SectionName = GetBookSectionName(bibleBookInfo.Name, i);

                if (string.IsNullOrEmpty(currentSectionGroupId))
                    currentSectionGroupId = AddTestamentSectionGroup(OldTestamentName);
                else if (i == OldTestamentBooksCount)
                    currentSectionGroupId = AddTestamentSectionGroup(NewTestamentName);

                var bookSectionId = AddBookSection(currentSectionGroupId, bibleBookInfo.SectionName, bibleBookInfo.Name);

                string bookFile = Path.Combine(ModuleFolder, bibleBookInfo.FileName);

                foreach (string line in File.ReadAllLines(bookFile, FileEncoding))
                {
                    string lineText = ShellText(line, moduleInfo); 

                    if (line.StartsWith(moduleInfo.ChapterSign))
                    {
                        if (currentChapterDoc != null)
                            UpdateChapterPage(currentChapterDoc);

                        if (ConvertChapterNameFunc != null)
                            lineText = ConvertChapterNameFunc(bibleBookInfo, lineText);

                        XmlNamespaceManager xnm;
                        currentChapterDoc = AddChapterPage(bookSectionId, lineText, 2, out xnm);

                        currentTableElement = AddTableToChapterPage(currentChapterDoc, xnm);
                    }
                    else if (line.StartsWith(moduleInfo.VerseSign))
                    {
                        if (currentTableElement == null)
                           throw new Exception("currentTableElement is null");

                        AddVerseRowToTable(currentTableElement, lineText);       
                    }
                }
            }

            if (currentChapterDoc != null)
            {
                UpdateChapterPage(currentChapterDoc);
            }
        }

        private string ShellText(string line, BibleQuotaModuleInfo moduleInfo)
        {
            var result = line.Replace("<<", "&lt;&lt;").Replace(">>", "&gt;&gt;");   // чтобы учитывать строки типа "<p>1 <<To the chief Musician on Neginoth, A Psalm of David.>> Hear me when I call, O God of my righteousness: thou hast enlarged me when I was in distress; have mercy upon me, and hear my prayer."

            result = StringUtils.GetText(result, moduleInfo.Alphabet).Trim();

            return result;
        }
    }
}
