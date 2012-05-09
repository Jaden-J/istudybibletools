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
        public List<string> Abbreviations { get; set; }
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
        

        public BibleQuotaConverter(string emptyNotebookName, string moduleFolder, string manifestFilePathToSave, Encoding fileEncoding,
            string oldTestamentName, string newTestamentName, int oldTestamentBooksCount, int newTestamentBooksCount, List<NotebookInfo> notebooksInfo)
            : base(emptyNotebookName, manifestFilePathToSave, oldTestamentName, newTestamentName, oldTestamentBooksCount, newTestamentBooksCount, notebooksInfo)
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
                        result.BibleBooksInfo[result.BibleBooksInfo.Count - 1].Abbreviations = value.ToLowerInvariant().Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).ToList();
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
                    string lineText = StringUtils.GetText(line, moduleInfo.Alphabet).Trim();

                    if (line.StartsWith(moduleInfo.ChapterSign))
                    {
                        if (currentChapterDoc != null)
                            oneNoteApp.UpdatePageContent(currentChapterDoc.ToString());

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
                oneNoteApp.UpdatePageContent(currentChapterDoc.ToString());
        }

        protected override void GenerateManifest(ExternalModuleInfo externalModuleInfo)
        {
            var extModuleInfo = (BibleQuotaModuleInfo)externalModuleInfo;

            ModuleInfo module = new ModuleInfo() { Name = extModuleInfo.Name, Version = "1.0", Notebooks = NotebooksInfo };
            module.BibleStructure = new BibleStructureInfo() 
            {
                Alphabet = extModuleInfo.Alphabet, 
                NewTestamentName = NewTestamentName, 
                OldTestamentName = OldTestamentName, 
                OldTestamentBooksCount = OldTestamentBooksCount,
                NewTestamentBooksCount = NewTestamentBooksCount,
                BibleBooks = new List<BibleBookInfo>() };

            foreach (var bibleBookInfo in extModuleInfo.BibleBooksInfo)
            {
                module.BibleStructure.BibleBooks.Add(new BibleBookInfo() { Name = bibleBookInfo.Name, SectionName = bibleBookInfo.SectionName, Abbreviations = bibleBookInfo.Abbreviations }); 
            }

            XmlSerializer ser = new XmlSerializer(typeof(ModuleInfo));
            using (var fs = new FileStream(ManifestFilePath, FileMode.Create))
            {
                ser.Serialize(fs, module);
                fs.Flush();
            }
        }
    }
}
