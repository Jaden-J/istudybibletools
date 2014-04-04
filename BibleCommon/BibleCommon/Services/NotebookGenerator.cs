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
using BibleCommon.Scheme;

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

        public static string AddSection(ref Application oneNoteApp, string sectionGroupId, string sectionName, bool addFirst = false)
        {
            XNamespace nms = XNamespace.Get(Constants.OneNoteXmlNs);
            XElement section = new XElement(nms + "Section", new XAttribute("name", sectionName));

            XmlNamespaceManager xnm;
            var sectionGroup = OneNoteUtils.GetHierarchyElement(ref oneNoteApp, sectionGroupId, HierarchyScope.hsSections, out xnm);

            if (addFirst)
                sectionGroup.Root.AddFirst(section);
            else
                sectionGroup.Root.Add(section);

            OneNoteUtils.UseOneNoteAPI(ref oneNoteApp, (oneNoteAppSafe) =>
            {
                oneNoteAppSafe.UpdateHierarchy(sectionGroup.ToString(), Constants.CurrentOneNoteSchema);
            });

            sectionGroup = OneNoteUtils.GetHierarchyElement(ref oneNoteApp, sectionGroupId, HierarchyScope.hsSections, out xnm);
            section = sectionGroup.Root.XPathSelectElement(string.Format("one:Section[@name=\"{0}\"]", sectionName), xnm);
            string sectionId = (string)section.Attribute("ID");

            return sectionId;
        }

        public static void DeleteHierarchy(ref Application oneNoteApp, string hierarchyId)
        {
            OneNoteUtils.UseOneNoteAPI(ref oneNoteApp, (oneNoteAppSafe) =>
            {
                oneNoteAppSafe.DeleteHierarchy(hierarchyId, default(DateTime), true);
            });           
        }


        public static XDocument AddPage(ref Application oneNoteApp, string sectionId, string pageTitle, int pageLevel, string locale, out XmlNamespaceManager xnm)
        {
            string pageId = null;

            OneNoteUtils.UseOneNoteAPI(ref oneNoteApp, (oneNoteAppSafe) =>
            {
                oneNoteAppSafe.CreateNewPage(sectionId, out pageId, NewPageStyle.npsBlankPageWithTitle);
            });

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

            OneNoteUtils.UpdatePageContentSafe(ref oneNoteApp, pageDocument, OneNoteUtils.GetOneNoteXNM());            

            var pageDoc = OneNoteUtils.GetPageContent(ref oneNoteApp, pageId, out xnm);

            return pageDoc;
        }

        public static void UpdatePageTitle(XDocument pageDoc, string newTitle, XmlNamespaceManager xnm)
        {
            var pageTitleEl = GetPageTitle(pageDoc, xnm);
            if (pageTitleEl != null)
                pageTitleEl.Value = newTitle;
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

        public static XElement AddTableToPage(XDocument pageDoc, bool bordersVisible, XmlNamespaceManager xnm, params CellInfo[] cells)
        {
            var nms = XNamespace.Get(Constants.OneNoteXmlNs);           

            var tableEl = GenerateTableElement(bordersVisible, xnm, cells);
            var oeEl = new XElement(nms + "OE", tableEl);
            OneNoteUtils.UpdateElementMetaData(oeEl, Constants.Key_Table, "true", xnm);

            pageDoc.Root.Add(new XElement(nms + "Outline",
                                            new XElement(nms + "OEChildren",
                                                oeEl
                                              )));

            return tableEl;
        }

        public static XElement GenerateTableElement(bool bordersVisible, XmlNamespaceManager xnm, params CellInfo[] cells)
        {
            var nms = XNamespace.Get(Constants.OneNoteXmlNs);           

            var columns = new XElement(nms + "Columns");                               

            for (int i = 0; i < cells.Length; i++)
            {
                columns.Add(new XElement(nms + "Column", new XAttribute("index", i), new XAttribute("width", cells[i].Width), new XAttribute("isLocked", true)));
            }

            var tableEl = new XElement(nms + "Table", new XAttribute("bordersVisible", bordersVisible), columns);            

            return tableEl;                
        }

        public static XElement GetPageTitle(XDocument pageDoc, XmlNamespaceManager xnm)
        {
            return pageDoc.Root.XPathSelectElement("one:Title/one:OE/one:T", xnm);
        }

        public static XElement GetPageTable(XDocument pageDoc, XmlNamespaceManager xnm)
        {
            var result = pageDoc.Root.XPathSelectElement(string.Format("//one:Outline/one:OEChildren/one:OE[./one:Meta[@name=\"{0}\"]]/one:Table", Constants.Key_Table), xnm);
            if (result == null)
                result = pageDoc.Root.XPathSelectElement("//one:Outline/one:OEChildren/one:OE/one:Table", xnm);

            return result;
        }

        public static XElement AddRowToTable(XElement tableElement, params XElement[] cells)
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

            return AddRowToTable(tableElement, cells.ToArray());
        }

        public static int AddColumnToTable(XElement tableElement, int cellWidth, XmlNamespaceManager xnm)
        {
            var nms = XNamespace.Get(Constants.OneNoteXmlNs);         

            var columnsEl = tableElement.XPathSelectElement("one:Columns", xnm);

            int columnsCount = columnsEl.Elements().Count();
            columnsEl.Add(new XElement(nms + "Column", new XAttribute("index", columnsCount), new XAttribute("width", cellWidth), new XAttribute("isLocked", true)));

            return columnsCount;
        }

        public static void RenameHierarchyElement(ref Application oneNoteApp, string hierarchyElementId, HierarchyScope scope, string newName)
        {
            XmlNamespaceManager xnm;
            var element = OneNoteUtils.GetHierarchyElement(ref oneNoteApp, hierarchyElementId, scope, out xnm);
            element.Root.SetAttributeValue("name", newName);

            OneNoteUtils.UseOneNoteAPI(ref oneNoteApp, (oneNoteAppSafe) =>
            {
                oneNoteAppSafe.UpdateHierarchy(element.ToString(), Constants.CurrentOneNoteSchema);
            });
        }

        public static XElement AddParallelVerseRowToBibleTable(XElement tableElement, SimpleVerse verse, int translationIndex, 
            SimpleVersePointer baseVerse, string locale, XmlNamespaceManager xnm)
        {
            var nms = XNamespace.Get(Constants.OneNoteXmlNs);

            var rows = tableElement.XPathSelectElements("one:Row", xnm);

            XElement verseRow = null;            
            foreach (var row in rows.Skip(1))
            {
                int rowChildsCount = row.Elements().Count();
                if (rowChildsCount <= translationIndex)
                {
                    if (translationIndex > 0)
                    {
                        var firstCell = row.XPathSelectElement("one:Cell[1]/one:OEChildren/one:OE/one:T", xnm);
                        if (firstCell != null)
                        {
                            var baseVerseNumber = VerseNumber.GetFromVerseText(firstCell.Value);
                            if (baseVerseNumber != baseVerse.VerseNumber)
                                continue;
                        }
                        verse.VerseLink = StringUtils.GetAttributeValue(firstCell.Value, "href");
                    }

                    verseRow = row;
                    break;
                }
            }    

            return AddParallelVerseCellToBibleRow(tableElement, verseRow, verse.GetVerseFullString(), translationIndex, locale);            
        }

        public static void AddParallelBibleTitle(XDocument pageDoc, XElement tableElement, string parallelTranslationModuleName, int bibleIndex, string locale, XmlNamespaceManager xnm)
        {
            var styleIndex = QuickStyleManager.AddQuickStyleDef(pageDoc, QuickStyleManager.StyleNameH2, QuickStyleManager.PredefinedStyles.H2, xnm);
            var cell = AddParallelVerseCellToBibleRow(tableElement, tableElement.XPathSelectElement("one:Row", xnm), parallelTranslationModuleName, bibleIndex, locale);
            QuickStyleManager.SetQuickStyleDefForCell(cell, styleIndex, xnm);            
        }

        public static XElement AddParallelVerseCellToBibleRow(XElement tableElement, XElement verseRow, string verseContent, int translationIndex, string locale)
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

            var cell = GetCell(verseContent, locale, nms);
            verseRow.Add(cell);

            return cell;
        }

        public static XElement GetCell(string locale, XNamespace nms, params XElement[] children)
        {
            var cell = new XElement(nms + "Cell",
                            new XElement(nms + "OEChildren",
                                children));

            if (!string.IsNullOrEmpty(locale))
                cell.Add(new XAttribute("lang", locale));

            return cell;
        }

        public static XElement[] TransformTextToParagraphs(string textToTransform, XNamespace nms)
        {
            var oeEls = new List<XElement>();

            foreach (var s in textToTransform.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries))
            {
                oeEls.Add(new XElement(nms + "OE", new XElement(nms + "T", new XCData(s))));                
            }

            return oeEls.ToArray();
        }

        public static void AddChildToCell(XElement cell, string content, XNamespace nms)
        {
            cell.Elements().First().Add(new XElement(nms + "OE",
                        new XElement(nms + "T",
                            new XCData(content))));
        }

        public static XElement GetCell(string cellText, string locale, XNamespace nms)
        {
            var cell = GetCell(locale, nms, new XElement(nms + "OE",                                    
                                  new XElement(nms + "T",
                                    new XCData(cellText))));

            return cell;
        }

        public static string GenerateBibleCommentsNotebook(ref Application oneNoteApp, string notebookName, 
            BibleStructureInfo bibleStructure, NotebooksStructure notebooksStructure, bool generateBookSectionGroups)
        {
            string targetNotebookId = NotebookGenerator.CreateNotebook(ref oneNoteApp, notebookName);

            var bookIndex = 0;

            foreach (var testament in notebooksStructure.Notebooks.First(n => n.Type == ContainerType.Bible).SectionGroups)
            {
                XElement testamentSectionGroupEl = AddRootSectionGroupToNotebook(ref oneNoteApp, targetNotebookId, testament.Name);

                if (generateBookSectionGroups)
                {
                    for (int i = 0; i < testament.SectionsCount; i++)
                    {
                        AddSectionGroup(ref oneNoteApp, testamentSectionGroupEl, bibleStructure.BibleBooks[bookIndex++].SectionName);
                    }
                }
            }

            return targetNotebookId;
        }

        public static string CreateNotebook(ref Application oneNoteApp, string notebookName)
        {
            string defaultNotebookFolderPath = null;

            OneNoteUtils.UseOneNoteAPI(ref oneNoteApp, (oneNoteAppSafe) =>
            {
                oneNoteAppSafe.GetSpecialLocation(SpecialLocation.slDefaultNotebookFolder, out defaultNotebookFolderPath);
            });

            var notebookId = CreateNotebook(ref oneNoteApp, notebookName, defaultNotebookFolderPath, null);

            if (!string.IsNullOrEmpty(notebookId))
            {
                OneNoteUtils.UseOneNoteAPI(ref oneNoteApp, (oneNoteAppSafe) =>
                {
                    oneNoteAppSafe.SyncHierarchy(notebookId);
                });
            }

            return notebookId;
        }

        public static void TryToRenameNotebookSafe(ref Application oneNoteApp, string notebookId, string notebookNickname)
        {
            try
            {
                XmlNamespaceManager xnm;
                var notebook = OneNoteUtils.GetHierarchyElement(ref oneNoteApp, notebookId, HierarchyScope.hsSelf, out xnm);

                notebook.Root.SetAttributeValue("nickname", notebookNickname);

                OneNoteUtils.UseOneNoteAPI(ref oneNoteApp, (oneNoteAppSafe) =>
                {
                    oneNoteAppSafe.UpdateHierarchy(notebook.ToString(), Constants.CurrentOneNoteSchema);
                });
            }
            catch (Exception ex)
            {
                Logger.LogError(ex);
            }
        }
        
        public static string CreateNotebook(ref Application oneNoteApp, string notebookName, string notebookDirectory, string nickname)
        {
            XmlNamespaceManager xnm;
            var nms = XNamespace.Get(Constants.OneNoteXmlNs);

            if (!Directory.Exists(notebookDirectory))
                Directory.CreateDirectory(notebookDirectory);

            var newNotebookPath = Utils.GetNewDirectoryPath(notebookDirectory + "\\" + notebookName);
            notebookName = Path.GetFileName(newNotebookPath);

            var notebookEl = new XElement(nms + "Notebook",
                                new XAttribute("name", notebookName),
                                new XAttribute("path", newNotebookPath));

            if (!string.IsNullOrEmpty(nickname))
                notebookEl.Add(new XAttribute("nickname", nickname));

            var notebooksEl = OneNoteUtils.GetHierarchyElement(ref oneNoteApp, null, HierarchyScope.hsNotebooks, out xnm);

            var lastNotebook = notebooksEl.Root.XPathSelectElements("one:Notebook", xnm).LastOrDefault();
            if (lastNotebook != null)
                lastNotebook.AddAfterSelf(notebookEl);
            else
                notebooksEl.Root.AddFirst(notebookEl);

            OneNoteUtils.UseOneNoteAPI(ref oneNoteApp, (oneNoteAppSafe) =>
            {
                oneNoteAppSafe.UpdateHierarchy(notebooksEl.ToString(), Constants.CurrentOneNoteSchema);
            });

            return OneNoteUtils.GetNotebookIdByName(ref oneNoteApp, notebookName, true);
        }        

        public static string GetBibleBookSectionName(string bookName, int bookIndex, int oldTestamentBooksCount)
        {
            int bookPrefix = bookIndex + 1 > oldTestamentBooksCount ? bookIndex + 1 - oldTestamentBooksCount : bookIndex + 1;
            return string.Format("{0:00}. {1}", bookPrefix, bookName);
        }

        public static XElement AddRootSectionGroupToNotebook(ref Application oneNoteApp, string notebookId, string sectionGroupName, string suffixIfSectionGroupExists = null)
        {
            XmlNamespaceManager xnm;
            var notebook = OneNoteUtils.GetHierarchyElement(ref oneNoteApp, notebookId, HierarchyScope.hsChildren, out xnm);

            if (notebook.Root.XPathSelectElement(string.Format("one:SectionGroup[@name=\"{0}\"]", sectionGroupName), xnm) != null)
            {
                if (!string.IsNullOrEmpty(suffixIfSectionGroupExists))  // иначе ошибку выдаст сам OneNote
                    sectionGroupName += suffixIfSectionGroupExists; 
            }

            AddSectionGroup(ref oneNoteApp, notebook.Root, sectionGroupName);

            notebook = OneNoteUtils.GetHierarchyElement(ref oneNoteApp, notebookId, HierarchyScope.hsChildren, out xnm);
            var newSectionGroup = notebook.Root.XPathSelectElement(string.Format("one:SectionGroup[@name=\"{0}\"]", sectionGroupName), xnm);
            return newSectionGroup;
        }       

        public static void AddSectionGroup(ref Application oneNoteApp, XElement parentElement, string sectionGroupName)
        {
            XNamespace nms = XNamespace.Get(Constants.OneNoteXmlNs);
            XElement newSectionGroup = new XElement(nms + "SectionGroup",
                                    new XAttribute("name", sectionGroupName));

            parentElement.Add(newSectionGroup);

            OneNoteUtils.UseOneNoteAPI(ref oneNoteApp, (oneNoteAppSafe) =>
            {
                oneNoteAppSafe.UpdateHierarchy(parentElement.ToString(), Constants.CurrentOneNoteSchema);
            });
        }     
    }
}
