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
using System.Xml.XPath;
using BibleCommon.Consts;
using BibleCommon.Handlers;
using BibleCommon.Common;

namespace BibleConfigurator.ModuleConverter
{
    public class DictionaryFile
    {
        public string TermPrefix { get; set; }
        public string FilePath { get; set; }
        public string DictionaryPageDescription { get; set; }
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

        private const int MaxTermsInPage = 100;                

        public enum StructureType
        {
            Strong,
            Dictionary
        }

        public Application _oneNoteApp;
        public bool IsOneNote2010 { get; set; }        
        public string NotebookId { get; set; }        
        public StructureType Type { get; set; }
        public string ManifestFilesFolder { get; set; }
        List<DictionaryFile> DictionaryFiles { get; set; }        
        public string Locale { get; set; }
        public Version Version { get; set; }
        public string TermStartString { get; set; }
        public string DictionaryModuleName { get; set; }
        public string DictionaryName { get; set; }
        public string DictionaryDescription { get; set; }        
        public string DictionarySectionGroupName { get; set; }
        public List<Exception> Errors { get; set; }
        public string UserNotesString { get; set; }
        public string FindAllVersesString { get; set; }
        public List<string> Terms { get; set; }
        public int PagesCount { get; set; }


        private string SectionGroupId { get; set; }
        private string SectionId  { get; set; }
        private Dictionary<string, int> LetterTermsCount { get; set; }
        private List<SectionInfo> DictionarySections { get; set; }
        private FindVersesWithStrongNumberHandler ProtocolHandler { get; set; }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="notebookName"></param>
        /// <param name="bqDictionaryFolder">Только один словарь может быть в папке.</param>
        /// <param name="type"></param>
        /// <param name="manifestFilesFolder"></param>
        public BibleQuotaDictionaryConverter(Application oneNoteApp, string notebookName, string dictionaryModuleName, string dictionaryName, string dictionaryDescription, 
            List<DictionaryFile> dictionaryFiles, StructureType type, string dictionarySectionGroupName, string manifestFilesFolder, string termStartString, string userNotesString, string findAllVersesString,
            string locale, Version version)
        {
            this.Type = type;
            this.DictionaryModuleName = dictionaryModuleName;
            this.DictionaryName = dictionaryName;
            this.DictionaryDescription = dictionaryDescription;            
            this.DictionarySectionGroupName = dictionarySectionGroupName;
            this._oneNoteApp = oneNoteApp;
            IsOneNote2010 = OneNoteUtils.IsOneNote2010Cached(oneNoteApp);
            this.ManifestFilesFolder = manifestFilesFolder;
            this.NotebookId = OneNoteUtils.GetNotebookIdByName(ref _oneNoteApp, notebookName, true);
            if (string.IsNullOrEmpty(this.NotebookId))
                this.NotebookId = NotebookGenerator.CreateNotebook(ref _oneNoteApp, notebookName);
            this.DictionaryFiles = dictionaryFiles;            
            this.Locale = locale;
            this.TermStartString = termStartString;            
            this.Errors = new List<Exception>();
            this.UserNotesString = userNotesString;
            this.FindAllVersesString = findAllVersesString;
            this.Version = version;
            this.Terms = new List<string>();

            if (!Directory.Exists(ManifestFilesFolder))
                Directory.CreateDirectory(ManifestFilesFolder);

            LetterTermsCount = new Dictionary<string, int>();
            DictionarySections = new List<SectionInfo>();
            ProtocolHandler = new FindVersesWithStrongNumberHandler();
        }        

        public void Convert()
        {
            var sectionGroupEl = NotebookGenerator.AddRootSectionGroupToNotebook(ref _oneNoteApp, NotebookId, this.DictionaryModuleName);
            SectionGroupId = (string)sectionGroupEl.Attribute("ID");
            XmlNamespaceManager xnm = OneNoteUtils.GetOneNoteXNM();

            foreach(var file in DictionaryFiles)
            {
                StringBuilder termDescription = null;
                string termName = null;
                int termsInPageCount = 0;
                int termIndex = file.StartIndex - 1;
                TermPageInfo pageInfo = null;

                if (Type == StructureType.Strong)
                {
                    SectionId = NotebookGenerator.AddSection(ref _oneNoteApp, SectionGroupId, Path.GetFileNameWithoutExtension(file.SectionName));
                    pageInfo = AddTermsPage(string.Format("{0:0000}-", file.StartIndex), file.DictionaryPageDescription);
                }
                                
                string prevTerm = null;
                foreach (string line in File.ReadAllLines(file.FilePath, Utils.GetFileEncoding(file.FilePath)))
                {
                    if (line.StartsWith(TermStartString))
                    {
                        if (!string.IsNullOrEmpty(termName))
                        {
                            termsInPageCount++;
                            termIndex++;

                            var newPageInfo = AddTermToPage(file, pageInfo, termName, termDescription.ToString(), termsInPageCount, prevTerm, ref termIndex, false, IsOneNote2010, xnm);
                            if (newPageInfo != null)
                            {
                                termsInPageCount = Type == StructureType.Strong ? 0 : 1;
                                pageInfo = newPageInfo;
                            }
                            prevTerm = termName;
                        }

                        termName = GetTermName(line, file);
                        termDescription = new StringBuilder();
                    }
                    else
                    {
                        if (termDescription != null)
                            termDescription.Append(ShellText(line));
                    }                    
                }

                if (!string.IsNullOrEmpty(termName))
                {                    
                    termIndex++;
                    AddTermToPage(file, pageInfo, termName, termDescription.ToString(), termsInPageCount, prevTerm, ref termIndex, true, IsOneNote2010, xnm);
                }
            }

            GenerateManifest();
        }     

        private TermPageInfo AddTermsPage(string pageName, string pageDisplayName)
        {
            XmlNamespaceManager xnm;
            var firstColumnWidthK = Type == StructureType.Strong ? 1 : 2;
            var pageDoc = NotebookGenerator.AddPage(ref _oneNoteApp, SectionId, pageName, 1, Locale, out xnm);
            var tableEl = NotebookGenerator.AddTableToPage(pageDoc, true, xnm, new CellInfo((int)(NotebookGenerator.MinimalCellWidth * firstColumnWidthK)), new CellInfo(SettingsManager.Instance.PageWidth_Bible));
            var styleIndex = QuickStyleManager.AddQuickStyleDef(pageDoc, QuickStyleManager.StyleNameH3, QuickStyleManager.PredefinedStyles.H3, xnm);
            if (!string.IsNullOrEmpty(pageDisplayName))
                AddPageTitle(pageDoc, pageDisplayName, xnm);

            OneNoteUtils.UpdatePageMetaData(pageDoc.Root, BibleCommon.Consts.Constants.Key_EmbeddedDictionaries,
                EmbeddedModuleInfo.Serialize(new List<EmbeddedModuleInfo>() { new EmbeddedModuleInfo(DictionaryModuleName, Version) }), xnm);

            PagesCount++;

            return new TermPageInfo() { PageDocument = pageDoc, TableElement = tableEl, StyleIndex = styleIndex };
        }

        private TermPageInfo AddTermToPage(DictionaryFile file, TermPageInfo pageInfo, string termName, string termDescription,
            int termsInPageCount, string prevTermName, ref int termIndex, bool isLatestTermInSection, bool isOneNote2010, XmlNamespaceManager xnm)
        {
            var nms = XNamespace.Get(Constants.OneNoteXmlNs);
            var pageInfoWasChanged = false;

            if (pageInfo == null || (Type == StructureType.Dictionary && !string.IsNullOrEmpty(prevTermName)
                && GetFirstTermValueChar(prevTermName) != GetFirstTermValueChar(termName)))
            {
                if (pageInfo != null)
                {
                    UpdateDictionaryTermsPage(pageInfo, prevTermName);
                }
                
                string sectionName = GetFirstTermValueChar(termName);
                SectionId = NotebookGenerator.AddSection(ref _oneNoteApp, SectionGroupId, sectionName);
                sectionName += ".one";
                if (!DictionarySections.Any(s => s.Name == sectionName))
                    DictionarySections.Add(new SectionInfo() { Name = sectionName});

                pageInfo = AddTermsPage(string.Format("{0:000}-", 1), file.DictionaryPageDescription);
                pageInfoWasChanged = true;
            }

            if (termDescription.StartsWith(Environment.NewLine))
                termDescription = termDescription.Remove(0, Environment.NewLine.Length);

            var termTable = NotebookGenerator.GenerateTableElement(false, new CellInfo(SettingsManager.Instance.PageWidth_Bible - 10));
            NotebookGenerator.AddRowToTable(termTable, NotebookGenerator.GetCell(termDescription, Locale, nms));
            var userNotesCell = NotebookGenerator.GetCell(UserNotesString, Locale, nms);
            QuickStyleManager.SetQuickStyleDefForCell(userNotesCell, pageInfo.StyleIndex, xnm);

            NotebookGenerator.AddRowToTable(termTable, userNotesCell);
            for (int i = 0; i <= 4; i++)
                NotebookGenerator.AddChildToCell(userNotesCell, string.Empty, nms);

            string termCellText;
            if (Type == StructureType.Strong)
            {                
                var commandUrl = ProtocolHandler.GetCommandUrl(termName);
                termCellText = string.Format("<b>{0}</b> <a href='{1}'><span style='font-size:8.0pt'>{2}</span></a>", termName, commandUrl, FindAllVersesString);
            }
            else
                termCellText = string.Format("<b>{0}</b>", termName);

            NotebookGenerator.AddRowToTable(pageInfo.TableElement,
                                NotebookGenerator.GetCell(termCellText, Locale, nms),
                                NotebookGenerator.GetCell(termTable, Locale, nms));

            Terms.Add(termName);
            if (Type == StructureType.Dictionary)
            {
                string firstChar = GetFirstTermValueChar(termName);
                if (!LetterTermsCount.ContainsKey(firstChar))
                    LetterTermsCount.Add(firstChar, 1);
                else
                    LetterTermsCount[firstChar]++;
            }

            if (Type == StructureType.Strong)
            {
                int bqTermIndex = int.Parse(StringUtils.GetText(termName.Substring(file.TermPrefix.Length)));
                if (bqTermIndex != termIndex)
                {
                    Errors.Add(new Exception(string.Format("bqTermIndex != termIndex: {0} != {1}", termName, termIndex)));
                    termIndex = bqTermIndex;
                }

                if ((termsInPageCount >= MaxTermsInPage && (termIndex % 100 == 99)) || isLatestTermInSection)
                {
                    string currentPageName = (string)pageInfo.PageDocument.Root.Attribute("name");
                    NotebookGenerator.UpdatePageTitle(pageInfo.PageDocument, currentPageName + termIndex.ToString("0000"), OneNoteUtils.GetOneNoteXNM());
                    UpdatePage(pageInfo.PageDocument);

                    if (!isLatestTermInSection)
                    {
                        pageInfo = AddTermsPage(string.Format("{0:0000}-", termIndex + 1), file.DictionaryPageDescription);
                        pageInfoWasChanged = true;
                    }
                }
            }
            else if (termsInPageCount >= MaxTermsInPage || isLatestTermInSection)
            {
                var letterTermsIndex = UpdateDictionaryTermsPage(pageInfo, termName);

                if (!isLatestTermInSection)
                {
                    pageInfo = AddTermsPage(string.Format("{0:000}-", letterTermsIndex + 1), file.DictionaryPageDescription);
                    pageInfoWasChanged = true;
                }
            }
            
            if (pageInfoWasChanged)
                return pageInfo;

            return null;
        }

        private int UpdateDictionaryTermsPage(TermPageInfo pageInfo, string termName)
        {
            var firstChar = GetFirstTermValueChar(termName);
            var letterTermsIndex = LetterTermsCount[firstChar];
            string currentPageName = (string)pageInfo.PageDocument.Root.Attribute("name");
            NotebookGenerator.UpdatePageTitle(pageInfo.PageDocument, currentPageName + letterTermsIndex.ToString("000"), OneNoteUtils.GetOneNoteXNM());
            UpdatePage(pageInfo.PageDocument);

            return letterTermsIndex;
        }

        private string GetFirstTermValueChar(string termName)
        {
            return char.ToUpper(termName[0]).ToString();            
        }

        private string GetTermName(string line, DictionaryFile file)
        {
            var result = StringUtils.GetText(line).Trim(new char[] { ' ', '"', '.' });
            if (Type == StructureType.Strong)
            {
                var number = int.Parse(result);
                result = string.Format("{0}{1:0000}", file.TermPrefix, number);
            }           

            return result;
        }

        private void AddPageTitle(XDocument pageDoc, string displayName, XmlNamespaceManager xnm)
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
                            .Replace("</p>", "</span>")
                            .Replace("<h6>", Environment.NewLine + "<b>")
                            .Replace("</h6>", "</b>")
                            .Replace("<h5>", Environment.NewLine + "<b>")
                            .Replace("</h5>", "</b>")
                            .Trim(new char[] { ' ', '\t' });                            
                                

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
            }

            //shell anchor
            result = ShellTag(result, "a");

            return result;
        }

        private static string ShellTag(string s, string tagName)
        {
            var startIndex = s.IndexOf(string.Format("<{0}", tagName));

            while(startIndex > -1)
            {
                var endIndex = s.IndexOf(string.Format("</{0}>", tagName), startIndex) + tagName.Length + 3;
                var text = StringUtils.GetText(s.Substring(startIndex, endIndex - startIndex));
                s = s.Substring(0, startIndex) + text + s.Substring(endIndex);                                
                startIndex = s.IndexOf(string.Format("<{0}", tagName), startIndex + text.Length);
            }

            return s;
        }

        protected void UpdatePage(XDocument pageDoc)
        {
            OneNoteUtils.UseOneNoteAPI(ref _oneNoteApp, () =>
            {
                _oneNoteApp.UpdatePageContent(pageDoc.ToString(), DateTime.MinValue, Constants.CurrentOneNoteSchema);
            });
        }

        private void GenerateManifest()
        {
            var module = new ModuleInfo()
            {
                ShortName = DictionaryModuleName,
                DisplayName = DictionaryName,
                Description = DictionaryDescription,
                Version = this.Version,
                Type = ModuleType.Dictionary,
                NotebooksStructure = new NotebooksStructure()
                {
                    DictionaryPagesCount = PagesCount,
                    DictionaryTermsCount = Terms.Count
                }
            };
            module.NotebooksStructure.Sections = Type == StructureType.Strong 
                ? this.DictionaryFiles.ConvertAll(df => new SectionInfo() { Name = df.SectionName })
                : DictionarySections;
            module.NotebooksStructure.DictionarySectionGroupName = this.DictionarySectionGroupName;

            Utils.SaveToXmlFile(
                module, 
                Path.Combine(ManifestFilesFolder, Constants.ManifestFileName));

            if (Type == StructureType.Dictionary)
            {
                Utils.SaveToXmlFile(
                    new ModuleDictionaryInfo() { TermSet = new TermSet() { Terms = this.Terms } },
                    Path.Combine(ManifestFilesFolder, Constants.DictionaryContentsFileName));
            }
        }

        public void Dispose()
        {
            _oneNoteApp = null;
        }
    }
}
