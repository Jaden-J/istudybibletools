﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using BibleCommon.Services;
using Microsoft.Office.Interop.OneNote;
using BibleCommon.Helpers;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;
using BibleCommon.Consts;
using BibleCommon.Handlers;

namespace BibleConfigurator.ModuleConverter
{
    public class DictionaryFile
    {
        public string TermPrefix { get; set; }
        public string FilePath { get; set; }
        public string DisplayName { get; set; }
        public string SectionName { get; set; }
        public int StartIndex { get; set; }        
    }

    public class BibleQuotaDictionaryConverter: IDisposable
    {
        internal class TermPageInfo
        {
            internal XDocument PageDocument { get; set; }
            internal XElement TableElement { get; set; }
            internal int StyleIndex { get; set; }
        }

        private const int MaxTermsInPageForStrong = 100;                

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
        public string DictionaryName { get; set; }        
        public List<Exception> Errors { get; set; }
        public string UserNotesString { get; set; }
        public string FindAllVersesString { get; set; }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="notebookName"></param>
        /// <param name="bqDictionaryFolder">Только один словарь может быть в папке.</param>
        /// <param name="type"></param>
        /// <param name="manifestFilesFolder"></param>
        public BibleQuotaDictionaryConverter(Application oneNoteApp, string notebookName, string dictionaryName,
            List<DictionaryFile> dictionaryFiles, StructureType type, string manifestFilesFolder, string termStartString, string userNotesString, string findAllVersesString,
            Encoding fileEncoding, string locale)
        {
            this.Type = type;
            this.DictionaryName = dictionaryName;
            this.OneNoteApp = oneNoteApp;
            this.ManifestFilesFolder = manifestFilesFolder;
            this.NotebookId = OneNoteUtils.GetNotebookIdByName(OneNoteApp, notebookName, true);
            if (string.IsNullOrEmpty(this.NotebookId))
                this.NotebookId = NotebookGenerator.CreateNotebook(OneNoteApp, notebookName);
            this.DictionaryFiles = dictionaryFiles;
            this.FileEncoding = fileEncoding;
            this.Locale = locale;
            this.TermStartString = termStartString;            
            this.Errors = new List<Exception>();
            this.UserNotesString = userNotesString;
            this.FindAllVersesString = findAllVersesString;
        }        

        public void Convert()
        {
            var sectionGroupEl = NotebookGenerator.AddRootSectionGroupToNotebook(OneNoteApp, NotebookId, this.DictionaryName);
            var sectionGroupId = (string)sectionGroupEl.Attribute("ID");
            XmlNamespaceManager xnm = OneNoteUtils.GetOneNoteXNM();

            foreach(var file in DictionaryFiles)
            {
                StringBuilder termDescription = null;
                string termName = null;

                var sectionId = NotebookGenerator.AddSection(OneNoteApp, sectionGroupId, file.SectionName);

                int termsInPageCount = 0;
                int termIndex = file.StartIndex - 1;
                var pageInfo = AddTermsPage(sectionId, Type == StructureType.Strong ? string.Format("{0:0000}-", file.StartIndex) : "First Alphabet Char", file.DisplayName);                                    

                foreach (string line in File.ReadAllLines(file.FilePath, FileEncoding))
                {
                    if (line.StartsWith(TermStartString))
                    {
                        if (!string.IsNullOrEmpty(termName))
                        {
                            termsInPageCount++;
                            termIndex++;

                            var newPageInfo = AddTermToPage(file, sectionId, pageInfo, termName, termDescription.ToString(), termsInPageCount, ref termIndex, false, xnm);
                            if (newPageInfo != null)
                            {
                                termsInPageCount = 0;
                                pageInfo = newPageInfo;
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
                    termIndex++;
                    AddTermToPage(file, sectionId, pageInfo, termName, termDescription.ToString(), termsInPageCount, ref termIndex, true, xnm);
                }
            }            
        }

        private TermPageInfo AddTermsPage(string sectionId, string pageName, string pageDisplayName)
        {
            XmlNamespaceManager xnm;
            var pageDoc = NotebookGenerator.AddPage(OneNoteApp, sectionId, pageName, 1, Locale, out xnm);
            var tableEl = NotebookGenerator.AddTableToPage(pageDoc, true, xnm, new CellInfo(NotebookGenerator.MinimalCellWidth), new CellInfo(SettingsManager.Instance.PageWidth_Bible));
            var styleIndex = QuickStyleManager.AddQuickStyleDef(pageDoc, QuickStyleManager.StyleNameH3, QuickStyleManager.PredefinedStyles.H3, xnm);
            if (!string.IsNullOrEmpty(pageDisplayName))
                AddPageDisplayName(pageDoc, pageDisplayName, xnm);
            return new TermPageInfo() { PageDocument = pageDoc, TableElement = tableEl, StyleIndex = styleIndex };
        }

        private TermPageInfo AddTermToPage(DictionaryFile file, string sectionId, TermPageInfo pageInfo, string termName, string termDescription,
            int termsInPageCount, ref int termIndex, bool isLatestTermInSection, XmlNamespaceManager xnm)
        {
            var nms = XNamespace.Get(Constants.OneNoteXmlNs);            
            
            var termTable = NotebookGenerator.GenerateTableElement(false, new CellInfo(SettingsManager.Instance.PageWidth_Bible - 10));
            NotebookGenerator.AddRowToTable(termTable, NotebookGenerator.GetCell(termDescription, Locale, nms));
            var userNotesCell = NotebookGenerator.GetCell(UserNotesString, Locale, nms);
            QuickStyleManager.SetQuickStyleDefForCell(userNotesCell, pageInfo.StyleIndex, xnm);

            NotebookGenerator.AddRowToTable(termTable, userNotesCell);
            for (int i = 0; i <= 4; i++)
                NotebookGenerator.AddChildToCell(userNotesCell, string.Empty, nms);

            var protocolHandler = new FindVersesWithStrongNumberHandler();
            var commandUrl = protocolHandler.GetCommandUrl(termName);

            NotebookGenerator.AddRowToTable(pageInfo.TableElement,
                                NotebookGenerator.GetCell(string.Format("<b>{0}</b> <a href='{1}'><span style='font-size:8.0pt'>{2}</span></a>", 
                                                            termName, commandUrl, FindAllVersesString), Locale, nms),
                                NotebookGenerator.GetCell(termTable, Locale, nms));

            if (Type == StructureType.Strong)
            {
                int bqTermIndex = int.Parse(StringUtils.GetText(termName.Substring(file.TermPrefix.Length)));
                if (bqTermIndex != termIndex)
                {
                    Errors.Add(new Exception(string.Format("bqTermIndex != termIndex: {0} != {1}", termName, termIndex)));
                    termIndex = bqTermIndex;
                }

                if ((termsInPageCount >= MaxTermsInPageForStrong && (termIndex % 100 == 99)) || isLatestTermInSection)
                {                    
                    string currentPageName = (string)pageInfo.PageDocument.Root.Attribute("name");
                    NotebookGenerator.UpdatePageTitle(pageInfo.PageDocument, currentPageName + termIndex.ToString("0000"), OneNoteUtils.GetOneNoteXNM());
                    UpdatePage(pageInfo.PageDocument);

                    if (!isLatestTermInSection)
                    {
                        return AddTermsPage(sectionId, string.Format("{0:0000}-", termIndex + 1), file.DisplayName);                                                
                    }
                }
            }

            return null;
        }      

        private string GetTermName(string line, DictionaryFile file)
        {
            var result = StringUtils.GetText(line).Trim();
            if (Type == StructureType.Strong)
            {
                var number = int.Parse(result);
                result = string.Format("{0}{1:0000}", file.TermPrefix, number);
            }

            return result;
        }

        private void AddPageDisplayName(XDocument pageDoc, string displayName, XmlNamespaceManager xnm)
        {
            var titleElement = pageDoc.Root.XPathSelectElement("one:Title", xnm);

            XNamespace nms = XNamespace.Get(Constants.OneNoteXmlNs);
            titleElement.AddAfterSelf(new XElement(nms + "Outline",
                             new XElement(nms + "Position",
                                    new XAttribute("x", "300.0"),
                                    new XAttribute("y", "14.40000057220459"),
                                    new XAttribute("z", "1")),
                             new XElement(nms + "OEChildren",
                               new XElement(nms + "OE",
                                   new XElement(nms + "T",
                                       new XCData(displayName))))));
        }

        private string ShellText(string text)
        {
            var result = text
                            .Replace("<br/>", Environment.NewLine)
                            .Replace("<br />", Environment.NewLine)
                            .Replace("<br>", Environment.NewLine)
                            .Replace("<p>", Environment.NewLine + "<span>")
                            .Replace("</p>", "</span>");

            if (Type == StructureType.Strong)
            {
                //shell font
                var indexOfFont = result.IndexOf("<font");
                if (indexOfFont != -1)
                {
                    var endIndexOfFontString = "</font>";
                    var endIndexOfFont = result.IndexOf(endIndexOfFontString, indexOfFont) + endIndexOfFontString.Length;
                    var fontFaceString = "face=\"";
                    var startIndex = result.IndexOf(fontFaceString, indexOfFont) + fontFaceString.Length;
                    var endIndex = result.IndexOf("\"", startIndex);
                    var fontName = result.Substring(startIndex, endIndex - startIndex);

                    result = result.Substring(0, indexOfFont)
                                + string.Format("<span style='font-family:{0}'>{1}</span>", fontName,
                                    StringUtils.GetText(result.Substring(indexOfFont, endIndexOfFont - indexOfFont)))
                                + result.Substring(endIndexOfFont);

                }

                //shell anchor
                result = ShellTag(result, "a");

            }

            return result;
        }

        private static string ShellTag(string s, string tagName)
        {
            var startIndex = s.IndexOf(string.Format("<{0}", tagName));
            if (startIndex != -1)
            {
                var endIndex = s.IndexOf(string.Format("</{0}>", tagName), startIndex) + tagName.Length + 3;
                s = s.Substring(0, startIndex) + StringUtils.GetText(s.Substring(startIndex, endIndex - startIndex)) + s.Substring(endIndex);
            }

            return s;
        }

        protected void UpdatePage(XDocument pageDoc)
        {
            OneNoteApp.UpdatePageContent(pageDoc.ToString(), DateTime.MinValue, Constants.CurrentOneNoteSchema);
        }

        public void Dispose()
        {
            OneNoteApp = null;
        }
    }
}
