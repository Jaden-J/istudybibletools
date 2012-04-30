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
        protected Application oneNoteApp { get; set; }
        protected string EmptyNotebookName { get; set; }
        protected string NotebookId { get; set; }

        public ConverterBase(string emptyNotebookName)
        {
            oneNoteApp = new Application();
            this.EmptyNotebookName = emptyNotebookName;
            this.NotebookId = OneNoteUtils.GetNotebookIdByName(oneNoteApp, EmptyNotebookName, true);
        }

        public void Convert()
        {
            var externalModuleInfo = ReadExternalModuleInfo();
            
         //   UpdateNotebookProperties(externalModuleInfo);            

            ProcessBibleBooks(externalModuleInfo);     
            
        }

        protected virtual void UpdateNotebookProperties(ExternalModuleInfo externalModuleInfo)
        {
            XmlNamespaceManager xnm;
            var notebook = OneNoteUtils.GetHierarchyElement(oneNoteApp, NotebookId, HierarchyScope.hsSelf, out xnm);

            notebook.Root.SetAttributeValue("name", externalModuleInfo.Name);
            notebook.Root.SetAttributeValue("nickname", externalModuleInfo.Name);
            oneNoteApp.UpdateHierarchy(notebook.ToString());
        }

        protected string AddNewBook(string bookName)
        {
            XmlNamespaceManager xnm;
            var notebook = OneNoteUtils.GetHierarchyElement(oneNoteApp, NotebookId, HierarchyScope.hsSections, out xnm);

            XNamespace nms = XNamespace.Get(Constants.OneNoteXmlNs);
            XElement section = new XElement(nms + "Section", new XAttribute("name", bookName));

            notebook.Root.Add(section);

            oneNoteApp.UpdateHierarchy(notebook.ToString());

            notebook = OneNoteUtils.GetHierarchyElement(oneNoteApp, NotebookId, HierarchyScope.hsSections, out xnm);

            var bookSection = notebook.Root.XPathSelectElement(string.Format("one:Section[@name='{0}']", bookName), xnm);
            return bookSection.Attribute("ID").Value;
        }

        protected XDocument AddNewChapter(string bookSectionId, string chapterName)
        {
            string pageId;
            oneNoteApp.CreateNewPage(bookSectionId, out pageId, NewPageStyle.npsBlankPageWithTitle);

            XNamespace nms = XNamespace.Get(Constants.OneNoteXmlNs);
            XDocument pageDocument = new XDocument(new XElement(nms + "Page",
                            new XAttribute("ID", pageId),                            
                            new XElement(nms + "Title",
                                new XElement(nms + "OE",
                                    new XElement(nms + "T",
                                        new XCData(
                                            chapterName
                                            ))))));

            oneNoteApp.UpdatePageContent(pageDocument.ToString());

            XmlNamespaceManager xnm;
            var pageDoc = OneNoteUtils.GetPageContent(oneNoteApp, pageId, out xnm);
            return pageDoc;
        }        

        protected abstract ExternalModuleInfo ReadExternalModuleInfo();
        protected abstract void ProcessBibleBooks(ExternalModuleInfo externalModuleInfo);
    }

}
