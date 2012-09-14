using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BibleCommon.Helpers;
using Microsoft.Office.Interop.OneNote;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;
using BibleCommon.Consts;
using System.Globalization;
using BibleCommon.Common;
using System.IO;

namespace BibleCommon.Services
{
    public class CellInfo
    {
        public int Width { get; set; }

        public CellInfo(int cellWidth)
        {
            this.Width = cellWidth;
        }
    }

    public static class NotebookGenerator
    {
        public const int MinimalCellWidth = 37;

        public static string AddSection(Application oneNoteApp, string sectionGroupId, string sectionName)
        {
            XNamespace nms = XNamespace.Get(Constants.OneNoteXmlNs);
            XElement section = new XElement(nms + "Section", new XAttribute("name", sectionName));

            XmlNamespaceManager xnm;
            var sectionGroup = OneNoteUtils.GetHierarchyElement(oneNoteApp, sectionGroupId, HierarchyScope.hsSections, out xnm);
            sectionGroup.Root.Add(section);
            oneNoteApp.UpdateHierarchy(sectionGroup.ToString(), Constants.CurrentOneNoteSchema);

            sectionGroup = OneNoteUtils.GetHierarchyElement(oneNoteApp, sectionGroupId, HierarchyScope.hsSections, out xnm);
            section = sectionGroup.Root.XPathSelectElement(string.Format("one:Section[@name='{0}']", sectionName), xnm);
            string sectionId = section.Attribute("ID").Value;

            return sectionId;
        }


        public static XDocument AddPage(Application oneNoteApp, string sectionId, string pageTitle, int pageLevel, string locale, out XmlNamespaceManager xnm)
        {
            string pageId;
            oneNoteApp.CreateNewPage(sectionId, out pageId, NewPageStyle.npsBlankPageWithTitle);

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

            if (!string.IsNullOrEmpty(locale))
                pageDocument.Root.Add(new XAttribute("lang", locale));

            oneNoteApp.UpdatePageContent(pageDocument.ToString(), DateTime.MinValue, Constants.CurrentOneNoteSchema);

            var pageDoc = OneNoteUtils.GetPageContent(oneNoteApp, pageId, out xnm);

            return pageDoc;
        }

        public static void AddTextElementToPage(XDocument pageDoc, string pageContent)
        {
            var nms = XNamespace.Get(Constants.OneNoteXmlNs);

            var textEl = new XElement(nms + "Outline",
                            new XElement(nms + "OEChildren",
                                new XElement(nms + "OE",
                                    new XElement(nms + "T",
                                        new XCData(
                                            pageContent
                                                    )))));

            pageDoc.Root.Add(textEl);
        }        

        public static XElement AddTableToPage(XDocument chapterDoc, bool bordersVisible, XmlNamespaceManager xnm, params CellInfo[] cells)
        {
            var nms = XNamespace.Get(Constants.OneNoteXmlNs);           

            var tableEl = new XElement(nms + "Outline",
                                            new XElement(nms + "OEChildren",
                                              new XElement(nms + "OE",
                                                  GenerateTableElement(bordersVisible, cells)
                                                  )));           

            chapterDoc.Root.Add(tableEl);

            return GetPageTable(chapterDoc, xnm);
        }

        public static XElement GenerateTableElement(bool bordersVisible, params CellInfo[] cells)
        {
            var nms = XNamespace.Get(Constants.OneNoteXmlNs);           

            var columns = new XElement(nms + "Columns");                               

            for (int i = 0; i < cells.Length; i++)
            {
                columns.Add(new XElement(nms + "Column", new XAttribute("index", i), new XAttribute("width", cells[i].Width), new XAttribute("isLocked", true)));
            }

            var tableEl = new XElement(new XElement(nms + "Table", new XAttribute("bordersVisible", bordersVisible), columns));

            return tableEl;                
        }

        public static XElement GetPageTable(XDocument chapterPageDoc, XmlNamespaceManager xnm)
        {
            return chapterPageDoc.Root.XPathSelectElement("//one:Outline/one:OEChildren/one:OE/one:Table", xnm);            
        }

        public static XElement AddRowToTable(XElement tableElement, List<XElement> cells)
        {
            var nms = XNamespace.Get(Constants.OneNoteXmlNs);           

            XElement newRow = new XElement(nms + "Row", cells);            

            tableElement.Add(newRow);

            return newRow;
        }

        public static XElement AddVerseRowToTable(XElement tableElement, string verseText, int emptyCellsCount, string locale)
        {
            var nms = XNamespace.Get(Constants.OneNoteXmlNs);

            var cells = new List<XElement>();

            cells.Add(GetCell(verseText, locale, nms));

            for (int i = 0; i < emptyCellsCount; i++)
            {
                cells.Add(GetCell(string.Empty, string.Empty, nms));
            }

            return AddRowToTable(tableElement, cells);
        }

        public static int AddColumnToTable(XElement tableElement, int cellWidth, XmlNamespaceManager xnm)
        {
            var nms = XNamespace.Get(Constants.OneNoteXmlNs);         

            var columnsEl = tableElement.XPathSelectElement("one:Columns", xnm);

            int columnsCount = columnsEl.Elements().Count();
            columnsEl.Add(new XElement(nms + "Column", new XAttribute("index", columnsCount), new XAttribute("width", cellWidth), new XAttribute("isLocked", true)));

            return columnsCount;
        }

        public static void RenameHierarchyElement(Application oneNoteApp, string hierarchyElementId, HierarchyScope scope, string newName)
        {
            XmlNamespaceManager xnm;
            var element = OneNoteUtils.GetHierarchyElement(oneNoteApp, hierarchyElementId, scope, out xnm);
            element.Root.SetAttributeValue("name", newName);
            oneNoteApp.UpdateHierarchy(element.ToString(), Constants.CurrentOneNoteSchema);
        }

        public static void AddParallelVerseRowToBibleTable(XElement tableElement, SimpleVerse verse, int translationIndex, 
            SimpleVersePointer baseVerse, string locale, XmlNamespaceManager xnm)
        {
            var nms = XNamespace.Get(Constants.OneNoteXmlNs);

            var rows = tableElement.XPathSelectElements("one:Row", xnm);

            XElement verseRow = null;
            int textBreakIndex, htmlBreakIndex;
            foreach (var row in rows.Skip(baseVerse.Verse - 1))
            {
                int rowChildsCount = row.Elements().Count();
                if (rowChildsCount <= translationIndex)
                {
                    if (translationIndex > 0)
                    {
                        var firstCell = row.XPathSelectElement("one:Cell[1]/one:OEChildren/one:OE/one:T", xnm);
                        if (firstCell != null)
                        {                            
                            var baseVerseNumber = StringUtils.GetNextString(firstCell.Value, -1, new SearchMissInfo(0, SearchMissInfo.MissMode.CancelOnMissFound), out textBreakIndex, out htmlBreakIndex);                                
                            if (baseVerseNumber != baseVerse.Verse.ToString())
                                continue;
                        }
                    }

                    verseRow = row;
                    break;
                }
            }

            AddParallelVerseCellToBibleRow(tableElement, verseRow, verse.VerseContent, translationIndex, locale);            
        }

        public static void AddParallelBibleTitle(XElement tableElement, string parallelTranslationModuleName, int bibleIndex, string locale, XmlNamespaceManager xnm)
        {
            AddParallelVerseCellToBibleRow(tableElement, tableElement.XPathSelectElement("one:Row", xnm), string.Format("<b>{0}</b>", parallelTranslationModuleName), bibleIndex, locale);            
        }

        public static void AddParallelVerseCellToBibleRow(XElement tableElement, XElement verseRow, string verseContent, int translationIndex, string locale)
        {
            var nms = XNamespace.Get(Constants.OneNoteXmlNs);

            if (verseRow == null)
            {   
                verseRow = new XElement(nms + "Row");

                for (int i = 0; i < translationIndex; i++)
                {
                    verseRow.Add(GetCell(string.Empty, string.Empty, nms));
                }
                
                tableElement.Add(verseRow);
            }

            verseRow.Add(GetCell(verseContent, locale, nms));
        }

        public static XElement GetCell(XElement child, string locale, XNamespace nms)
        {
            var cell = new XElement(nms + "Cell",
                            new XElement(nms + "OEChildren",
                                new XElement(nms + "OE",
                                    child
                                    )));

            if (!string.IsNullOrEmpty(locale))
                cell.Add(new XAttribute("lang", locale));

            return cell;
        }

        public static XElement GetCell(string cellText, string locale, XNamespace nms)
        {
            var cell = GetCell(new XElement(nms + "T",
                                    new XCData(cellText)), locale, nms);

            return cell;
        }

        public static void GenerateSummaryOfNotesNotebook(Application oneNoteApp, string bibleNotebookName, string targetEmptyNotebookName)
        {
            string bibleNotebookId = OneNoteUtils.GetNotebookIdByName(oneNoteApp, bibleNotebookName, false);
            XmlNamespaceManager xnm;
            var bibleNotebookDoc = OneNoteUtils.GetHierarchyElement(oneNoteApp, bibleNotebookId, HierarchyScope.hsSections, out xnm);

            string targetNotebookId = OneNoteUtils.GetNotebookIdByName(oneNoteApp, targetEmptyNotebookName, false);                        

            foreach (var testamentSectionGroup in bibleNotebookDoc.Root.XPathSelectElements("one:SectionGroup", xnm))
            {
                string testamentSectionGroupName = testamentSectionGroup.Attribute("name").Value;
                XElement testamentSectionGroupEl =  AddRootSectionGroupToNotebook(oneNoteApp, targetNotebookId, testamentSectionGroupName);                

                foreach (var bibleBookSection in testamentSectionGroup.XPathSelectElements("one:Section", xnm))
                {
                    string bibleBookSectionName = bibleBookSection.Attribute("name").Value;
                    AddSectionGroup(oneNoteApp, testamentSectionGroupEl, bibleBookSectionName);
                }
            }
        }

        public static string CreateNotebook(Application oneNoteApp, string notebookName)
        {
            string defaultNotebookFolderPath;
            oneNoteApp.GetSpecialLocation(SpecialLocation.slDefaultNotebookFolder, out defaultNotebookFolderPath);
            return CreateNotebook(oneNoteApp, notebookName, defaultNotebookFolderPath);
        }

        public static string CreateNotebook(Application oneNoteApp, string notebookName, string notebookDirectory)
        {
            XmlNamespaceManager xnm;
            var nms = XNamespace.Get(Constants.OneNoteXmlNs);

            var newNotebookPath = Utils.GetNewDirectoryPath(notebookDirectory + "\\" + notebookName);
            notebookName = Path.GetFileName(newNotebookPath);

            var notebookEl = new XElement(nms + "Notebook",
                                new XAttribute("name", notebookName),
                                new XAttribute("path", newNotebookPath));
            var notebooksEl = OneNoteUtils.GetHierarchyElement(oneNoteApp, null, HierarchyScope.hsNotebooks, out xnm);
            notebooksEl.Root.Elements().Last().AddBeforeSelf(notebookEl);
            oneNoteApp.UpdateHierarchy(notebooksEl.ToString(), Constants.CurrentOneNoteSchema);

            return OneNoteUtils.GetNotebookIdByName(oneNoteApp, notebookName, true);
        }        

        public static string GetBibleBookSectionName(string bookName, int bookIndex, int oldTestamentBooksCount)
        {
            int bookPrefix = bookIndex + 1 > oldTestamentBooksCount ? bookIndex + 1 - oldTestamentBooksCount : bookIndex + 1;
            return string.Format("{0:00}. {1}", bookPrefix, bookName);
        }

        public static XElement AddRootSectionGroupToNotebook(Application oneNoteApp, string notebookId, string sectionGroupName)
        {
            XmlNamespaceManager xnm;
            var notebook = OneNoteUtils.GetHierarchyElement(oneNoteApp, notebookId, HierarchyScope.hsChildren, out xnm);

            AddSectionGroup(oneNoteApp, notebook.Root, sectionGroupName);

            notebook = OneNoteUtils.GetHierarchyElement(oneNoteApp, notebookId, HierarchyScope.hsChildren, out xnm);
            var newSectionGroup = notebook.Root.XPathSelectElement(string.Format("one:SectionGroup[@name='{0}']", sectionGroupName), xnm);
            return newSectionGroup;
        }

        public static void AddSectionGroup(Application oneNoteApp, XElement parentElement, string sectionGroupName)
        {
            XNamespace nms = XNamespace.Get(Constants.OneNoteXmlNs);
            XElement newSectionGroup = new XElement(nms + "SectionGroup",
                                    new XAttribute("name", sectionGroupName));

            parentElement.Add(newSectionGroup);
            oneNoteApp.UpdateHierarchy(parentElement.ToString(), Constants.CurrentOneNoteSchema);
        }     
    }
}
