﻿using System;
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

namespace BibleCommon.Services
{
    public static class NotebookGenerator
    {
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="chapterDoc"></param>
        /// <param name="bibleCellWidth"></param>
        /// <param name="bibleIndex">для параллельных переводов - какая по счёту Библия</param>
        /// <param name="xnm"></param>
        /// <returns></returns>
        public static XElement AddTableToBibleChapterPage(XDocument chapterDoc, int bibleCellWidth, XmlNamespaceManager xnm)
        {
            var nms = XNamespace.Get(Constants.OneNoteXmlNs);           

            var tableEl = new XElement(nms + "Outline",
                                            new XElement(nms + "OEChildren",
                                              new XElement(nms + "OE",
                                                  new XElement(nms + "Table", new XAttribute("bordersVisible", false),
                                                      new XElement(nms + "Columns",
                                                          new XElement(nms + "Column", new XAttribute("index", 0), new XAttribute("width", bibleCellWidth), new XAttribute("isLocked", true)),
                                                          new XElement(nms + "Column", new XAttribute("index", 1), new XAttribute("width", 37), new XAttribute("isLocked", true))
                                                              )))));

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

        public static void AddVerseRowToBibleTable(XElement tableElement, string verseText, string locale)
        {
            var nms = XNamespace.Get(Constants.OneNoteXmlNs);

            var cell1 = GetCell(verseText, nms);

            if (!string.IsNullOrEmpty(locale))
                cell1.Add(new XAttribute("lang", locale));

            var cell2 = GetCell(string.Empty, nms);

            if (!string.IsNullOrEmpty(locale))
                cell2.Add(new XAttribute("lang", locale));

            XElement newRow = new XElement(nms + "Row", cell1, cell2);

            tableElement.Add(newRow);
        }

        public static int ExtendBibleTableForParallelTranslation(XElement tableElement, int bibleCellWidth, XmlNamespaceManager xnm)
        {
            var nms = XNamespace.Get(Constants.OneNoteXmlNs);         

            var columnsEl = tableElement.XPathSelectElement("one:Columns", xnm);

            int columnsCount = columnsEl.Elements().Count();
            columnsEl.Add(new XElement(nms + "Column", new XAttribute("index", columnsCount), new XAttribute("width", bibleCellWidth), new XAttribute("isLocked", true)));

            return columnsCount;
        }

        public static void AddParallelVerseRowToBibleTable(XElement tableElement, SimpleVerse verse, int translationIndex, string locale, XmlNamespaceManager xnm)
        {
            var nms = XNamespace.Get(Constants.OneNoteXmlNs);

            var rows = tableElement.XPathSelectElements("one:Row", xnm);

            XElement verseRow = null;
            
            foreach (var row in rows.Skip(verse.Verse - 1))
            {
                int rowChildsCount = row.Elements().Count();
                if (rowChildsCount <= translationIndex)
                {
                    if (translationIndex > 0)
                    {
                        var firstCell = row.XPathSelectElement("one:Cell[1]/one:OEChildren/one:OE/one:T", xnm);
                        if (firstCell != null)
                        {
                            int? index = StringUtils.GetStringFirstNumber(firstCell.Value);
                            if (!index.HasValue || index.Value != verse.Verse)
                                continue;
                        }
                    }

                    verseRow = row;
                    break;
                }
            }

            if (verseRow == null)
            {
                verseRow = new XElement(nms + "Row");

                for (int i = 0; i < translationIndex; i++)
                {
                    verseRow.Add(GetCell(string.Empty, nms));
                }
                
                tableElement.Add(verseRow);
            }

            verseRow.Add(GetCell(verse.VerseContent, nms));
        }

        private static XElement GetCell(string cellText, XNamespace nms)
        {
            return new XElement(nms + "Cell",
                            new XElement(nms + "OEChildren",
                                new XElement(nms + "OE",
                                    new XElement(nms + "T",
                                        new XCData(
                                            cellText
                                                    )))));
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
                XElement testamentSectionGroupEl =  OneNoteUtils.AddRootSectionGroupToNotebook(oneNoteApp, targetNotebookId, testamentSectionGroupName);                

                foreach (var bibleBookSection in testamentSectionGroup.XPathSelectElements("one:Section", xnm))
                {
                    string bibleBookSectionName = bibleBookSection.Attribute("name").Value;
                    OneNoteUtils.AddSectionGroup(oneNoteApp, testamentSectionGroupEl, bibleBookSectionName);
                }
            }
        }
    }
}
