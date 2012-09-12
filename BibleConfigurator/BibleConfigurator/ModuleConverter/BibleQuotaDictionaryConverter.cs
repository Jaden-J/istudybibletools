using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using BibleCommon.Services;
using Microsoft.Office.Interop.OneNote;
using BibleCommon.Helpers;
using System.Xml;
using System.Xml.Linq;
using BibleCommon.Consts;

namespace BibleConfigurator.ModuleConverter
{
    public class DictionaryFile
    {
        public string TermPrefix { get; set; }
        public string FilePath { get; set; }
        public string SectionGroupName { get; set; }
        public int StartIndex { get; set; }
    }

    public class BibleQuotaDictionaryConverter: IDisposable
    {
        private const int maxFileInSectionForStrong = 100;                

        public enum StructureType
        {
            Strong,
            Dictionary
        }

        public Application OneNoteApp { get; set; }
        public string NotebookId { get; set; }        
        public StructureType Type { get; set; }
        public string ManifestFilesFolder { get; set; }
        List<DictionaryFile> DictionaryFiles { get; set; }
        public Encoding FileEncoding { get; set; }
        public string Locale { get; set; }
        public string Version { get; set; }
        public string TermStartString { get; set; }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="newNotebookName"></param>
        /// <param name="bqDictionaryFolder">Только один словарь может быть в папке.</param>
        /// <param name="type"></param>
        /// <param name="manifestFilesFolder"></param>
        public BibleQuotaDictionaryConverter(Application oneNoteApp, string newNotebookName, 
            List<DictionaryFile> dictionaryFiles, StructureType type, string manifestFilesFolder, string termStartString,
            Encoding fileEncoding, string locale, string version)
        {
            this.Type = type;
            this.OneNoteApp = oneNoteApp;
            this.ManifestFilesFolder = manifestFilesFolder;
            this.NotebookId = NotebookGenerator.CreateNotebook(OneNoteApp, newNotebookName);
            this.DictionaryFiles = dictionaryFiles;
            this.FileEncoding = fileEncoding;
            this.Locale = locale;
            this.TermStartString = termStartString;
        }

        private string GetPageName(string termName)
        {
            return string.Format("{0}", termName);
        }

        public void Convert()
        {            
            foreach(var file in DictionaryFiles)
            {
                StringBuilder termDescription = null;
                string termName = null;

                var sectionGroupEl = NotebookGenerator.AddRootSectionGroupToNotebook(OneNoteApp, NotebookId, file.SectionGroupName);
                var sectionGroupId = sectionGroupEl.Attribute("ID").Value;

                int pagesInSectionCount = 0;
                int pageIndex = file.StartIndex - 1;                
                var sectionId = Type == StructureType.Strong 
                                    ? NotebookGenerator.AddSection(OneNoteApp, sectionGroupId, string.Format("{0:0000}-", file.StartIndex)) 
                                    : "First Alphabet char";                

                foreach (string line in File.ReadAllLines(file.FilePath, FileEncoding))
                {
                    if (line.StartsWith(TermStartString))
                    {
                        if (!string.IsNullOrEmpty(termName))
                        {
                            pagesInSectionCount++;
                            pageIndex++;

                            var newSectionId = AddTermPage(file, sectionGroupId, sectionId, termName, termDescription.ToString(), pagesInSectionCount, pageIndex);
                            if (!string.IsNullOrEmpty(newSectionId))
                            {
                                pagesInSectionCount = 0;
                                sectionId = newSectionId;
                            }
                        }

                        termName = GetTermName(line, file);
                        termDescription = new StringBuilder();
                    }
                    else
                    {
                        termDescription.Append(ShellText(line));
                    }                    
                }

                if (!string.IsNullOrEmpty(termName))
                {
                    AddTermPage(file, sectionGroupId, sectionId, termName, termDescription.ToString(), pagesInSectionCount, pageIndex);
                }
            }            
        }

        private string GetTermName(string line, DictionaryFile file)
        {
            var result = StringUtils.GetText(line);
            if (Type == StructureType.Strong)
            {
                result = result.Substring(result.Length - 4, 4);                
                result = file.TermPrefix + result;
            }

            return result;
        }
        

        private string ShellText(string text)
        {
            var result = text.Replace("<br/>", Environment.NewLine).Replace("<br />", Environment.NewLine).Replace("<br>", Environment.NewLine);

            if (Type == StructureType.Strong)
            {                
                var indexOfFont = result.IndexOf("<font");
                if (indexOfFont != -1)
                {
                    var endIndexOfFontString = "</font>";
                    var endIndexOfFont = result.IndexOf(endIndexOfFontString, indexOfFont) + endIndexOfFontString.Length;
                    var fontFaceString = "face=\"";
                    var startIndex = result.IndexOf(fontFaceString) + fontFaceString.Length;
                    var endIndex = result.IndexOf("\"", startIndex);
                    var fontName = result.Substring(startIndex, endIndex - startIndex);

                    result = result.Substring(0, indexOfFont)
                                + string.Format("<span style='font-family:{0}'>{1}</span>", fontName, 
                                    StringUtils.GetText(result.Substring(indexOfFont, endIndexOfFont - indexOfFont)))
                                + result.Substring(endIndexOfFont);

                }                
            }

            return result;
        }        
        
        /// <summary>
        /// 
        /// </summary>
        /// <param name="file"></param>
        /// <param name="sectionGroupId"></param>
        /// <param name="sectionId"></param>
        /// <param name="termName"></param>
        /// <param name="termDescription"></param>
        /// <param name="pagesInSectionCount"></param>
        /// <param name="pageIndex"></param>
        /// <returns>new section Id</returns>
        private string AddTermPage(DictionaryFile file, string sectionGroupId, string sectionId, string termName, string termDescription, int pagesInSectionCount, int pageIndex)
        {
            XmlNamespaceManager xnm;
            var termPageDoc = NotebookGenerator.AddPage(OneNoteApp, sectionId, GetPageName(termName), 1, Locale, out xnm);
            NotebookGenerator.AddTextElementToPage(termPageDoc, termDescription);

            UpdatePage(termPageDoc);

            if (Type == StructureType.Strong)
            {
                if (int.Parse(StringUtils.GetText(termName.Substring(file.TermPrefix.Length))) != pageIndex)
                    throw new Exception(string.Format("termName != fileIndex: {0} != {1}", termName, pageIndex));

                if (pagesInSectionCount >= maxFileInSectionForStrong)
                {
                    string latestSectionName = OneNoteUtils.GetHierarchyElementName(OneNoteApp, sectionId);
                    NotebookGenerator.RenameHierarchyElement(OneNoteApp, sectionId, HierarchyScope.hsSections, latestSectionName + pageIndex.ToString("0000"));                    
                    sectionId = NotebookGenerator.AddSection(OneNoteApp, sectionGroupId, string.Format("{0:0000}-", pageIndex + 1));
                    return sectionId;
                }
            }

            return string.Empty;
        }

        protected virtual void UpdatePage(XDocument pageDoc)
        {
            OneNoteApp.UpdatePageContent(pageDoc.ToString(), DateTime.MinValue, Constants.CurrentOneNoteSchema);
        }

        public void Dispose()
        {
            OneNoteApp = null;
        }
    }
}
