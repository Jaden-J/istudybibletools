﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BibleCommon.Common;
using System.Xml.Linq;
using BibleCommon.Helpers;
using Microsoft.Office.Interop.OneNote;
using System.Xml.XPath;
using System.Xml;
using BibleCommon.Consts;

namespace BibleCommon.Services
{
    public static class NotesPageManager
    {
        public static string UpdateNotesPage(Application oneNoteApp, NoteLinkManager noteLinkManager, VersePointer vp, bool isChapter,
           HierarchySearchManager.HierarchyObjectInfo verseHierarchyObjectInfo,
           PageIdInfo notePageId, string notesPageId, string notePageContentObjectId,
           string notesPageName, int notesPageWidth, bool force)
        {
            string targetContentObjectId = string.Empty;
            XNamespace nms = XNamespace.Get(Constants.OneNoteXmlNs);
            OneNoteProxy.PageContent notesPageDocument = OneNoteProxy.Instance.GetPageContent(oneNoteApp, notesPageId, OneNoteProxy.PageType.NotesPage);

            XElement rowElement = GetNotesRowAndCreateIfNotExists(oneNoteApp, vp, isChapter, notesPageWidth, verseHierarchyObjectInfo,
                notesPageDocument.Content, notesPageDocument.Xnm, nms);

            if (rowElement != null)
            {
                AddLinkToNotesPage(oneNoteApp, noteLinkManager, vp, rowElement, notePageId, 
                    notePageContentObjectId, notesPageDocument, notesPageDocument.Xnm, nms, notesPageName, force);

                targetContentObjectId = GetNotesRowObjectId(oneNoteApp, notesPageId, vp.Verse, isChapter);
            }

            return targetContentObjectId;
        }

        private static void AddLinkToNotesPage(Application oneNoteApp, NoteLinkManager noteLinkManager, VersePointer vp, XElement rowElement,
           PageIdInfo notePageId, string notePageContentObjectId,
           OneNoteProxy.PageContent notesPageDocument, XmlNamespaceManager xnm, XNamespace nms, string notesPageName, bool force)
        {
            string noteTitle = (notePageId.SectionGroupName != notePageId.SectionName && !string.IsNullOrEmpty(notePageId.SectionGroupName))
                ? string.Format("{0} / {1} / {2}", notePageId.SectionGroupName, notePageId.SectionName, notePageId.PageName)
                : string.Format("{0} / {1}", notePageId.SectionName, notePageId.PageName);

            XElement suchNoteLink = null;
            XElement notesCellElement = rowElement.XPathSelectElement("one:Cell[2]/one:OEChildren", xnm);

            string link = OneNoteUtils.GenerateHref(oneNoteApp, noteTitle, notePageId.PageId, notePageContentObjectId);
            int pageIdStringIndex = link.IndexOf("page-id={");
            if (pageIdStringIndex != -1)
            {
                string pageId = link.Substring(pageIdStringIndex, link.IndexOf('}', pageIdStringIndex) - pageIdStringIndex + 1);
                suchNoteLink = rowElement.XPathSelectElement(string.Format(
                   "one:Cell[2]/one:OEChildren/one:OE/one:T[contains(.,'{0}')]", pageId), xnm);
            }

            if (suchNoteLink != null)
            {
                var key = new NoteLinkManager.NotePageProcessedVerseId() { NotePageId = notePageId.PageId, NotesPageName = notesPageName };
                if (force && !noteLinkManager.ContainsNotePageProcessedVerse(key, vp))  // если в первый раз и force                
                {  // удаляем старые ссылки на текущую странцу, так как мы начали новый анализ с параметром "force" и мы только в первый раз зашли сюда
                    var verseLinks = suchNoteLink.Parent.NextNode;
                    if (verseLinks != null && verseLinks.XPathSelectElement("one:List", xnm) == null)
                        verseLinks.Remove();

                    suchNoteLink.Parent.Remove();
                    suchNoteLink = null;
                }
            }

            if (suchNoteLink != null)
                OneNoteUtils.NormalizeTextElement(suchNoteLink);

            if (suchNoteLink == null)  // если нет ссылки на такую же заметку
            {
                XNode prevLink = null;
                foreach (XElement existingLink in rowElement.XPathSelectElements("one:Cell[2]/one:OEChildren/one:OE/one:T", xnm))
                {
                    if (existingLink.Parent.XPathSelectElement("one:List", xnm) != null)  // если мы смотрим ссылку с номером, а не строку типа "ссылка1; ссылка2"
                    {
                        string existingNoteTitle = StringUtils.GetText(existingLink.Value);

                        if (noteTitle.CompareTo(existingNoteTitle) < 0)
                            break;
                        prevLink = existingLink.Parent;
                    }
                }

                XElement linkElement = new XElement(nms + "OE",
                                            new XElement(nms + "List",
                                                        new XElement(nms + "Number", new XAttribute("numberSequence", 0), new XAttribute("numberFormat", "##."))),
                                            new XElement(nms + "T",
                                                new XCData(
                                                    link + GetMultiVerseString(vp.ParentVersePointer ?? vp))));

                if (prevLink == null)
                {
                    notesCellElement.AddFirst(linkElement);
                }
                else
                {
                    if (prevLink.NextNode != null && prevLink.NextNode.XPathSelectElement("one:List", xnm) == null)  // если следующая строка типа "ссылка1; ссылка2"                    
                        prevLink = prevLink.NextNode;

                    prevLink.AddAfterSelf(linkElement);
                }
            }
            else
            {
                string pageLink = OneNoteUtils.GenerateHref(oneNoteApp, noteTitle, notePageId.PageId, notePageId.PageTitleId);

                var verseLinksOE = suchNoteLink.Parent.NextNode;
                if (verseLinksOE != null && verseLinksOE.XPathSelectElement("one:List", xnm) == null)  // значит следующая строка без номера, то есть значит идут ссылки
                {
                    XElement existingVerseLinksElement = verseLinksOE.XPathSelectElement("one:T", xnm);


                    int currentVerseIndex = existingVerseLinksElement.Value.Split(new string[] { "</a>" }, StringSplitOptions.None).Length;

                    existingVerseLinksElement.Value += Resources.Constants.VerseLinksDelimiter + OneNoteUtils.GenerateHref(oneNoteApp,
                                string.Format(Resources.Constants.VerseLinkTemplate, currentVerseIndex), notePageId.PageId, notePageContentObjectId)
                                + GetMultiVerseString(vp.ParentVersePointer ?? vp);

                }
                else  // значит мы нашли второе упоминание данной ссылки в заметке
                {
                    string firstVerseLink = StringUtils.GetAttributeValue(suchNoteLink.Value, "href");
                    firstVerseLink = string.Format("<a href='{0}'>{1}</a>", firstVerseLink, string.Format(Resources.Constants.VerseLinkTemplate, 1));
                    XElement verseLinksElement = new XElement(nms + "OE",
                                                    new XElement(nms + "T",
                                                        new XCData(StringUtils.MultiplyString("&nbsp;", 8) +
                                                            string.Join(Resources.Constants.VerseLinksDelimiter, new string[] { 
                                                                firstVerseLink + GetExistingMultiVerseString(suchNoteLink), 
                                                                OneNoteUtils.GenerateHref(oneNoteApp, 
                                                                    string.Format(Resources.Constants.VerseLinkTemplate, 2), notePageId.PageId, notePageContentObjectId)
                                                                    + GetMultiVerseString(vp.ParentVersePointer ?? vp) })
                                                            )));

                    suchNoteLink.Parent.AddAfterSelf(verseLinksElement);
                }

                suchNoteLink.Value = pageLink;

                if (suchNoteLink.Parent.XPathSelectElement("one:List", xnm) == null)  // почему то нет номера у строки
                    suchNoteLink.Parent.AddFirst(new XElement(nms + "List",
                                                    new XElement(nms + "Number", new XAttribute("numberSequence", 0), new XAttribute("numberFormat", "##."))));

            }

            notesPageDocument.WasModified = true;
        }      

       

        private static string GetMultiVerseString(VersePointer vp)
        {
            if (vp.IsMultiVerse)
            {
                if (vp.TopChapter != null && vp.TopVerse != null)
                    return string.Format(" <b>({0}:{1}-{2}:{3})</b>", vp.Chapter, vp.Verse, vp.TopChapter, vp.TopVerse);
                else if (vp.TopChapter != null && vp.IsChapter)
                    return string.Format(" <b>({0}-{1})</b>", vp.Chapter, vp.TopChapter);
                else
                    return string.Format(" <b>(:{0}-{1})</b>", vp.Verse, vp.TopVerse);
            }
            else
                return string.Empty;
        }

        private static string GetExistingMultiVerseString(XElement suchNoteLink)
        {
            string multiVerseString = string.Empty;

            string topVerseSearchPattern = "(:";
            string topVerseEndSearchPattern = ")";

            int topVerseIndex = -1;
            string suchNoteLinkText = string.Empty;

            if (suchNoteLink != null)
                suchNoteLinkText = StringUtils.GetText(suchNoteLink.Value);

            if (!string.IsNullOrEmpty(suchNoteLinkText))
                topVerseIndex = suchNoteLinkText.IndexOf(topVerseSearchPattern);

            if (topVerseIndex != -1)
            {
                int topVerseEndIndex = suchNoteLinkText.IndexOf(topVerseEndSearchPattern, topVerseIndex + 1);
                if (topVerseEndIndex != -1)
                {
                    multiVerseString = suchNoteLinkText.Substring(topVerseIndex, topVerseEndIndex - topVerseIndex + 1);
                }
            }

            if (!string.IsNullOrEmpty(multiVerseString))
                return string.Format(" <b>{0}</b>", multiVerseString);

            return multiVerseString;
        }

        internal static string GetNotesRowObjectId(Application oneNoteApp, string notesPageId, int? verseNumber, bool isChapter)
        {
            string result = string.Empty;
            OneNoteProxy.PageContent notesPageDocument = OneNoteProxy.Instance.GetPageContent(oneNoteApp, notesPageId, OneNoteProxy.PageType.NotesPage);
            XElement tableElement = notesPageDocument.Content.XPathSelectElement("//one:Outline/one:OEChildren/one:OE/one:Table", notesPageDocument.Xnm);
            XElement targetElement = GetNotesRow(tableElement, verseNumber, isChapter, notesPageDocument.Xnm);

            if (targetElement != null)
                result = (string)targetElement.XPathSelectElement("one:Cell/one:OEChildren/one:OE", notesPageDocument.Xnm).Attribute("objectID");

            return result;
        }

        private static XElement GetNotesRowAndCreateIfNotExists(Application oneNoteApp, VersePointer vp, bool isChapter, int mainColumnWidth, HierarchySearchManager.HierarchyObjectInfo verseHierarchyObjectInfo,
            XDocument notesPageDocument, XmlNamespaceManager xnm, XNamespace nms)
        {
            XElement rootElement = notesPageDocument.XPathSelectElement("//one:Outline/one:OEChildren/one:OE", xnm);
            if (rootElement == null)
            {
                notesPageDocument.Root.Add(new XElement(nms + "Outline",
                                              new XElement(nms + "OEChildren",
                                                new XElement(nms + "OE",
                                                    new XElement(nms + "Table", new XAttribute("bordersVisible", true),
                                                        new XElement(nms + "Columns",
                                                            new XElement(nms + "Column", new XAttribute("index", 0), new XAttribute("width", 37), new XAttribute("isLocked", true)),
                                                            new XElement(nms + "Column", new XAttribute("index", 1), new XAttribute("width", mainColumnWidth), new XAttribute("isLocked", true))
                                                                ))))));
                rootElement = notesPageDocument.XPathSelectElement("//one:Outline/one:OEChildren/one:OE", xnm);
            }

            XElement tableElement = rootElement.XPathSelectElement("one:Table", xnm);

            if (tableElement == null)
            {
                rootElement.Add(new XElement(nms + "Table", new XAttribute("bordersVisible", true)));

                tableElement = rootElement.XPathSelectElement("one:Table", xnm);
            }

            XElement rowElement = GetNotesRow(tableElement, vp.Verse, isChapter, xnm);

            if (rowElement == null)
            {
                AddNewNotesRow(oneNoteApp, vp, isChapter, verseHierarchyObjectInfo, tableElement, xnm, nms);

                rowElement = GetNotesRow(tableElement, vp.Verse, isChapter, xnm);
            }

            return rowElement;
        }

        private static XElement GetNotesRow(XElement tableElement, int? verseNumber, bool isChapter, XmlNamespaceManager xnm)
        {

            XElement result = !isChapter ?
                                tableElement
                                   .XPathSelectElement(string.Format("one:Row/one:Cell[1]/one:OEChildren/one:OE/one:T[contains(.,'>:{0}<')]", verseNumber.GetValueOrDefault(0)), xnm)
                              : tableElement
                                   .XPathSelectElement("one:Row/one:Cell[1]/one:OEChildren/one:OE/one:T[normalize-space(.)='']", xnm)
                                ;

            if (result != null)
                result = result.Parent.Parent.Parent.Parent;

            return result;
        }

        private static void AddNewNotesRow(Application oneNoteApp, VersePointer vp, bool isChapter, HierarchySearchManager.HierarchyObjectInfo verseHierarchyObjectInfo,
            XElement tableElement, XmlNamespaceManager xnm, XNamespace nms)
        {
            XElement newRow = new XElement(nms + "Row",
                                    new XElement(nms + "Cell",
                                        new XElement(nms + "OEChildren",
                                            new XElement(nms + "OE",
                                                new XElement(nms + "T",
                                                    new XCData(
                                                        !isChapter ?
                                                            OneNoteUtils.GenerateHref(oneNoteApp, string.Format(":{0}", vp.Verse.GetValueOrDefault(0)),
                                                                verseHierarchyObjectInfo.PageId, verseHierarchyObjectInfo.ContentObjectId)
                                                            :
                                                            string.Empty
                                                                ))))),
                                    new XElement(nms + "Cell",
                                        new XElement(nms + "OEChildren")));

            XElement prevRow = null;

            if (!isChapter)  // иначе добавляем первым
            {
                foreach (var row in tableElement.XPathSelectElements("one:Row/one:Cell[1]/one:OEChildren/one:OE/one:T", xnm))
                {
                    XText verseData = (XText)row.Nodes().First();
                    int? verseNumber = StringUtils.GetStringLastNumber(verseData.Value);
                    if (verseNumber.GetValueOrDefault(0) > vp.Verse)
                        break;

                    prevRow = row.Parent.Parent.Parent.Parent;
                }
            }

            if (prevRow == null)
                prevRow = tableElement.XPathSelectElement("one:Columns", xnm);

            if (prevRow == null)
                tableElement.AddFirst(newRow);
            else
                prevRow.AddAfterSelf(newRow);
        }


    }
}