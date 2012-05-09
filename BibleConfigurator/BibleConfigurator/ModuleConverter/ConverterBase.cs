using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.OneNote;
using BibleCommon.Helpers;
using BibleCommon.Services;
using System.Xml.Linq;
using BibleCommon.Consts;
using System.Xml;
using System.Xml.XPath;
using BibleCommon.Common;
using System.IO;

namespace BibleConfigurator.ModuleConverter
{
    public class ExternalModuleInfo
    {
        public string Name { get; set; }
        public string ShortName { get; set; }
        public string Alphabet { get; set; }
        public int BooksCount { get; set; }        
    }

    

    public abstract class ConverterBase
    {
        protected abstract ExternalModuleInfo ReadExternalModuleInfo();
        protected abstract void ProcessBibleBooks(ExternalModuleInfo externalModuleInfo);
        protected abstract void GenerateManifest(ExternalModuleInfo externalModuleInfo);

        protected Application oneNoteApp { get; set; }
        protected string EmptyNotebookName { get; set; }
        protected string NotebookId { get; set; }
        protected string ManifestFilePath { get; set; }
        protected string OldTestamentName { get; set; }
        protected string NewTestamentName { get; set; }        
        protected int OldTestamentBooksCount { get; set; }
        protected int NewTestamentBooksCount { get; set; }
        protected List<NotebookInfo> NotebooksInfo { get; set; }

        public ConverterBase(string emptyNotebookName, string manifestFilePathToSave,
            string oldTestamentName, string newTestamentName, int oldTestamentBooksCount, int newTestamentBooksCount, List<NotebookInfo> notebooksInfo)
        {
            oneNoteApp = new Application();
            this.EmptyNotebookName = emptyNotebookName;
            this.NotebookId = OneNoteUtils.GetNotebookIdByName(oneNoteApp, EmptyNotebookName, true);
            this.ManifestFilePath = manifestFilePathToSave;
            this.OldTestamentName = oldTestamentName;
            this.NewTestamentName = newTestamentName;
            this.OldTestamentBooksCount = oldTestamentBooksCount;
            this.NewTestamentBooksCount = newTestamentBooksCount;
            this.NotebooksInfo = notebooksInfo;           
        }

        public void Convert()
        {
            var externalModuleInfo = ReadExternalModuleInfo();
            
            UpdateNotebookProperties(externalModuleInfo);            

            ProcessBibleBooks(externalModuleInfo);

            GenerateManifest(externalModuleInfo);
        }

        protected virtual string GetBookSectionName(string bookName, int bookIndex)
        {
            int bookPrefix = bookIndex + 1 > OldTestamentBooksCount ? bookIndex + 1 - OldTestamentBooksCount : bookIndex + 1;
            return string.Format("{0:00}. {1}", bookPrefix, bookName);
        }

        protected virtual void UpdateNotebookProperties(ExternalModuleInfo externalModuleInfo)
        {
            XmlNamespaceManager xnm;
            var notebook = OneNoteUtils.GetHierarchyElement(oneNoteApp, NotebookId, HierarchyScope.hsSelf, out xnm);

            string notebookName = Path.GetFileNameWithoutExtension(NotebooksInfo.First(n => n.Type == NotebookType.Bible).Name);

            notebook.Root.SetAttributeValue("name", notebookName);
            notebook.Root.SetAttributeValue("nickname", notebookName);

            oneNoteApp.UpdateHierarchy(notebook.ToString());
        }

        protected virtual string AddTestamentSectionGroup(string testmanetName)
        {
            XNamespace nms = XNamespace.Get(Constants.OneNoteXmlNs);
            XElement testamentSectionGroup = new XElement(nms + "SectionGroup",
                                    new XAttribute("name", testmanetName));

            XmlNamespaceManager xnm;
            var notebook = OneNoteUtils.GetHierarchyElement(oneNoteApp, NotebookId, HierarchyScope.hsChildren, out xnm);
            notebook.Root.Add(testamentSectionGroup);
            oneNoteApp.UpdateHierarchy(notebook.ToString());

            notebook = OneNoteUtils.GetHierarchyElement(oneNoteApp, NotebookId, HierarchyScope.hsChildren, out xnm);
            testamentSectionGroup = notebook.Root.XPathSelectElement(string.Format("one:SectionGroup[@name='{0}']", testmanetName), xnm);
            return testamentSectionGroup.Attribute("ID").Value;              
        }

        protected virtual string AddBookSection(string sectionGroupId, string sectionName, string bookName)
        {
            XNamespace nms = XNamespace.Get(Constants.OneNoteXmlNs);
            XElement section = new XElement(nms + "Section", new XAttribute("name", sectionName));

            XmlNamespaceManager xnm;
            var sectionGroup = OneNoteUtils.GetHierarchyElement(oneNoteApp, sectionGroupId, HierarchyScope.hsSections, out xnm);
            sectionGroup.Root.Add(section);
            oneNoteApp.UpdateHierarchy(sectionGroup.ToString());

            sectionGroup = OneNoteUtils.GetHierarchyElement(oneNoteApp, sectionGroupId, HierarchyScope.hsSections, out xnm);
            section = sectionGroup.Root.XPathSelectElement(string.Format("one:Section[@name='{0}']", sectionName), xnm);
            string sectionId = section.Attribute("ID").Value;

            AddChapterPage(sectionId, bookName, 1, out xnm);

            return sectionId;
        }

        protected virtual XDocument AddChapterPage(string bookSectionId, string pageTitle, int pageLevel, out XmlNamespaceManager xnm)
        {
            string pageId;
            oneNoteApp.CreateNewPage(bookSectionId, out pageId, NewPageStyle.npsBlankPageWithTitle);

            var nms = XNamespace.Get(Constants.OneNoteXmlNs);
            XDocument pageDocument = new XDocument(new XElement(nms + "Page",
                            new XAttribute("ID", pageId),
                            new XAttribute("pageLevel", pageLevel),
                            new XElement(nms + "Title",
                                new XElement(nms + "OE",
                                    new XElement(nms + "T",
                                        new XCData(
                                            pageTitle
                                            ))))));

            oneNoteApp.UpdatePageContent(pageDocument.ToString());
            
            var pageDoc = OneNoteUtils.GetPageContent(oneNoteApp, pageId, out xnm);
            return pageDoc;
        }

        protected virtual XElement AddTableToChapterPage(XDocument chapterDoc, XmlNamespaceManager xnm)
        {
            var nms = XNamespace.Get(Constants.OneNoteXmlNs);

            chapterDoc.Root.Add(new XElement(nms + "Outline",
                                            new XElement(nms + "OEChildren",
                                              new XElement(nms + "OE",
                                                  new XElement(nms + "Table", new XAttribute("bordersVisible", false),
                                                      new XElement(nms + "Columns",
                                                          new XElement(nms + "Column", new XAttribute("index", 0), new XAttribute("width", 500), new XAttribute("isLocked", true)),
                                                          new XElement(nms + "Column", new XAttribute("index", 1), new XAttribute("width", 37), new XAttribute("isLocked", true))
                                                              ))))));

            return chapterDoc.Root.XPathSelectElement("//one:Outline/one:OEChildren/one:OE/one:Table", xnm);
        }

        protected virtual void AddVerseRowToTable(XElement tableElement, string verseText)
        {
            var nms = XNamespace.Get(Constants.OneNoteXmlNs);

            XElement newRow = new XElement(nms + "Row",
                                  new XElement(nms + "Cell",
                                      new XElement(nms + "OEChildren",
                                          new XElement(nms + "OE",
                                              new XElement(nms + "T",
                                                  new XCData(
                                                      verseText
                                                              ))))),
                                  new XElement(nms + "Cell",
                                      new XElement(nms + "OEChildren",
                                          new XElement(nms + "OE",
                                              new XElement(nms + "T",
                                                  new XCData(
                                                      string.Empty
                                                              ))))));

            tableElement.Add(newRow);                 
        }
    }
}
