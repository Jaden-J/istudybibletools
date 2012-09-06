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
    public static class NotebookGenerator
    {
        public const int MinimalCellWidth = 37;

        public static string AddBookSectionToBibleNotebook(Application oneNoteApp, string sectionGroupId, string sectionName, string bookName)
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


        public static XDocument AddChapterPageToBibleNotebook(Application oneNoteApp, string bookSectionId, string pageTitle, int pageLevel, string locale, out XmlNamespaceManager xnm)
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

            if (!string.IsNullOrEmpty(locale))
                pageDocument.Root.Add(new XAttribute("lang", locale));

            oneNoteApp.UpdatePageContent(pageDocument.ToString(), DateTime.MinValue, Constants.CurrentOneNoteSchema);

            var pageDoc = OneNoteUtils.GetPageContent(oneNoteApp, pageId, out xnm);

            return pageDoc;
        }
        
        public static XElement AddTableToBibleChapterPage(XDocument chapterDoc, int bibleCellWidth, int columnsCount, XmlNamespaceManager xnm)
        {
            var nms = XNamespace.Get(Constants.OneNoteXmlNs);

            var columns = new XElement(nms + "Columns", 
                                new XElement(nms + "Column", new XAttribute("index", 0), new XAttribute("width", bibleCellWidth), new XAttribute("isLocked", true)));

            for (int i = 0; i < columnsCount - 1; i++)
            {
                columns.Add(new XElement(nms + "Column", new XAttribute("index", i + 1), new XAttribute("width", MinimalCellWidth), new XAttribute("isLocked", true)));
            }

            var tableEl = new XElement(nms + "Outline",
                                            new XElement(nms + "OEChildren",
                                              new XElement(nms + "OE",
                                                  new XElement(nms + "Table", new XAttribute("bordersVisible", false),
                                                      columns 
                                                    ))));

            //var outlines = chapterDoc.Root.XPathSelectElements("//one:Outline", xnm);
            //int bibleIndex = outlines.Count();
            //if (bibleIndex > 0)
            //{
            //    var lastOutline = outlines.Last();

            //    if (lastOutline != null)
            //    {
            //        var prevPosition = lastOutline.XPathSelectElement("one:Position", xnm);
            //        var prevX = prevPosition.Attribute("x");
            //        var prevWidth = lastOutline.XPathSelectElement("one:Size", xnm).Attribute("width");

            //        var newX = (prevX != null ? float.Parse(prevX.Value, CultureInfo.InvariantCulture) : 0) + (prevWidth != null ? float.Parse(prevWidth.Value, CultureInfo.InvariantCulture) : 0) + 30;

            //        var positionEl = new XElement(nms + "Position",
            //                            new XAttribute("x", newX),
            //                            new XAttribute("y", prevPosition.Attribute("y").Value),
            //                            new XAttribute("z", prevPosition.Attribute("z").Value));

            //        tableEl.AddFirst(positionEl);
            //    }
            //}

            chapterDoc.Root.Add(tableEl);

            return GetBibleTable(chapterDoc, xnm);
        }

        public static XElement GetBibleTable(XDocument chapterPageDoc, XmlNamespaceManager xnm)
        {
            return chapterPageDoc.Root.XPathSelectElement("//one:Outline/one:OEChildren/one:OE/one:Table", xnm);            
        }

        public static void AddVerseRowToBibleTable(XElement tableElement, string verseText, int emptyCellsCount, string locale)
        {
            var nms = XNamespace.Get(Constants.OneNoteXmlNs);

            var cell1 = GetCell(verseText, locale, nms);                        

            XElement newRow = new XElement(nms + "Row", cell1);

            for (int i = 0; i < emptyCellsCount; i++)
            {
                newRow.Add(GetCell(string.Empty, string.Empty, nms));
            }

            tableElement.Add(newRow);
        }

        public static int AddColumnToTable(XElement tableElement, int cellWidth, XmlNamespaceManager xnm)
        {
            var nms = XNamespace.Get(Constants.OneNoteXmlNs);         

            var columnsEl = tableElement.XPathSelectElement("one:Columns", xnm);

            int columnsCount = columnsEl.Elements().Count();
            columnsEl.Add(new XElement(nms + "Column", new XAttribute("index", columnsCount), new XAttribute("width", cellWidth), new XAttribute("isLocked", true)));

            return columnsCount;
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

        public static XElement GetCell(string cellText, string locale, XNamespace nms)
        {
            var cell = new XElement(nms + "Cell",
                            new XElement(nms + "OEChildren",
                                new XElement(nms + "OE",
                                    new XElement(nms + "T",
                                        new XCData(
                                            cellText
                                                    )))));

            if (!string.IsNullOrEmpty(locale))
                cell.Add(new XAttribute("lang", locale));

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
            XmlNamespaceManager xnm;
            var nms = XNamespace.Get(Constants.OneNoteXmlNs);
            string defaultNotebookFolderPath;

            oneNoteApp.GetSpecialLocation(SpecialLocation.slDefaultNotebookFolder, out defaultNotebookFolderPath);
            var newNotebookPath = Utils.GetNewDirectoryPath(defaultNotebookFolderPath + "\\" + notebookName);
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
