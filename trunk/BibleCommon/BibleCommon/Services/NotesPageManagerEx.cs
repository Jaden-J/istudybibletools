using System;
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
using System.Text.RegularExpressions;
using BibleCommon.Contracts;

namespace BibleCommon.Services
{
    public class NotesPageManagerEx : INotesPageManager
    {
        public static readonly string Const_ManagerName = "NotesPageManagerEx";

        private XNamespace _nms;        

        public string ManagerName
        {
            get { return Const_ManagerName; }
        }

        public NotesPageManagerEx()
        {
            _nms = XNamespace.Get(Constants.OneNoteXmlNs);            
        }

        public string UpdateNotesPage(ref Application oneNoteApp, NoteLinkManager noteLinkManager, VersePointer vp, bool isChapter,
           HierarchySearchManager.HierarchyObjectInfo verseHierarchyObjectInfo,
           HierarchyElementInfo notePageId, string notesPageId, string notePageContentObjectId,
           string notesPageName, int notesPageWidth, bool force, bool processAsExtendedVerse, out bool rowWasAdded)
        {
            string targetContentObjectId = string.Empty;            
            
            OneNoteProxy.PageContent notesPageDocument = OneNoteProxy.Instance.GetPageContent(ref oneNoteApp, notesPageId, OneNoteProxy.PageType.NotesPage);            

            var rootElement = GetRootElementAndCreateIfNotExists(ref oneNoteApp, vp, isChapter, notesPageWidth, verseHierarchyObjectInfo,
                notesPageDocument, out rowWasAdded);

            if (rootElement != null)
            {
                AddLinkToNotesPage(ref oneNoteApp, noteLinkManager, vp, rootElement, notePageId,
                    notePageContentObjectId, notesPageDocument, notesPageName, force, processAsExtendedVerse);

                targetContentObjectId = GetNotesRowObjectId(ref oneNoteApp, notesPageId, verseHierarchyObjectInfo.VerseNumber, isChapter);
            }

            return targetContentObjectId;
        }


         private XElement GetRootElementAndCreateIfNotExists(ref Application oneNoteApp, VersePointer vp, bool isChapter,
            int mainColumnWidth, HierarchySearchManager.HierarchyObjectInfo verseHierarchyObjectInfo,
            OneNoteProxy.PageContent notesPageDocument, out bool rowWasAdded)
        {
            rowWasAdded = false;

            XElement rootElement = notesPageDocument.Content.XPathSelectElement("//one:Outline/one:OEChildren", notesPageDocument.Xnm);
            if (rootElement == null)
            {
                notesPageDocument.Content.Root.Add(new XElement(_nms + "Outline",
                                                        new XElement(_nms + "OEChildren",
                                                            new XAttribute("indent", 2)
                                                    )));
                rootElement = notesPageDocument.Content.XPathSelectElement("//one:Outline/one:OEChildren", notesPageDocument.Xnm);
            }

            return rootElement;
        }

        public string GetNotesRowObjectId(ref Application oneNoteApp, string notesPageId, VerseNumber? verseNumber, bool isChapter)
        {
            string result = string.Empty;
            OneNoteProxy.PageContent notesPageDocument = OneNoteProxy.Instance.GetPageContent(ref oneNoteApp, notesPageId, OneNoteProxy.PageType.NotesPage);
            XElement tableElement = notesPageDocument.Content.XPathSelectElement("//one:Outline/one:OEChildren/one:OE/one:Table", notesPageDocument.Xnm);
            XElement targetElement = GetNotesRow(tableElement, verseNumber, isChapter, notesPageDocument.Xnm);

            if (targetElement != null)
                result = (string)targetElement.XPathSelectElement("one:Cell/one:OEChildren/one:OE", notesPageDocument.Xnm).Attribute("objectID");

            return result;
        }

        private XElement _parentElement;

        private void AddLinkToNotesPage(ref Application oneNoteApp, NoteLinkManager noteLinkManager, VersePointer vp, XElement rootElement,
           HierarchyElementInfo notePageInfo, string notePageContentObjectId,
           OneNoteProxy.PageContent notesPageDocument, string notesPageName, bool force, bool processAsExtendedVerse)
        {
            _parentElement = rootElement;
            if (notePageInfo.Parent != null)
                CreateParentTreeStructure(notePageInfo.Parent, notesPageDocument.Xnm);

            string link = OneNoteUtils.GenerateHref(ref oneNoteApp, notePageInfo.Name, notePageInfo.Id, notePageContentObjectId);

            var suchNoteLink = SearchExistingNoteLink(notesPageDocument, notePageInfo, link);

            if (suchNoteLink != null)
            {
                var key = new NoteLinkManager.NotePageProcessedVerseId() { NotePageId = notePageInfo.Id, NotesPageName = notesPageName };
                if (force && !noteLinkManager.ContainsNotePageProcessedVerse(key, vp) && !processAsExtendedVerse)  // если в первый раз и force и не расширенный стих
                {  // удаляем старые ссылки на текущую странцу, так как мы начали новый анализ с параметром "force" и мы только в первый раз зашли сюда
                    RemoveExistingNoteLink(suchNoteLink, notesPageDocument);                    
                }
            }

            if (suchNoteLink != null)
                OneNoteUtils.NormalizeTextElement(suchNoteLink);

            if (suchNoteLink == null)  // если нет ссылки на такую же заметку
            {
                //XNode prevLink = null;
                //foreach (XElement existingLink in rowElement.XPathSelectElements("one:Cell[2]/one:OEChildren/one:OE/one:T", xnm))
                //{
                //    if (existingLink.Parent.XPathSelectElement("one:List", xnm) != null)  // если мы смотрим ссылку с номером, а не строку типа "ссылка1; ссылка2"
                //    {
                //        string existingNoteTitle = StringUtils.GetText(existingLink.Value);

                //        if (noteTitle.CompareTo(existingNoteTitle) < 0)
                //            break;
                //        prevLink = existingLink.Parent;
                //    }
                //}

                XElement linkElement = new XElement(_nms + "OE",
                                            new XElement(_nms + "List",
                                                        new XElement(_nms + "Number", new XAttribute("numberSequence", 0), new XAttribute("numberFormat", "##."))),
                                            new XElement(_nms + "T",
                                                new XCData(
                                                    link + GetMultiVerseString(vp.ParentVersePointer ?? vp))));
                OneNoteUtils.UpdateElementMetaData(linkElement, Constants.Key_Id, notePageInfo.Id, notesPageDocument.Xnm);

                _parentElement.Add(linkElement);

                //if (prevLink == null)
                //{
                //    notesCellElement.AddFirst(linkElement);
                //}
                //else
                //{
                //    if (prevLink.NextNode != null && prevLink.NextNode.XPathSelectElement("one:List", notesPageDocument.Xnm) == null)  // если следующая строка типа "ссылка1; ссылка2"                    
                //        prevLink = prevLink.NextNode;
                //    prevLink.AddAfterSelf(linkElement);
                //}
            }
            else if (!processAsExtendedVerse)
            {
                string pageLink = OneNoteUtils.GenerateHref(ref oneNoteApp, notePageInfo.Name, notePageInfo.Id, notePageInfo.PageTitleId);

                var verseLinksOE = suchNoteLink.Parent.NextNode;
                if (verseLinksOE != null && verseLinksOE.XPathSelectElement("one:List", notesPageDocument.Xnm) == null)  // значит следующая строка без номера, то есть значит идут ссылки
                {
                    XElement existingVerseLinksElement = verseLinksOE.XPathSelectElement("one:T", notesPageDocument.Xnm);

                    int currentVerseIndex = existingVerseLinksElement.Value.Split(new string[] { "</a>" }, StringSplitOptions.None).Length;

                    existingVerseLinksElement.Value += Resources.Constants.VerseLinksDelimiter + OneNoteUtils.GenerateHref(ref oneNoteApp,
                                string.Format(Resources.Constants.VerseLinkTemplate, currentVerseIndex), notePageInfo.Id, notePageContentObjectId)
                                + GetMultiVerseString(vp.ParentVersePointer ?? vp);

                }
                else  // значит мы нашли второе упоминание данной ссылки в заметке
                {
                    string firstVerseLink = StringUtils.GetAttributeValue(suchNoteLink.Value, "href");
                    firstVerseLink = string.Format("<a href='{0}'>{1}</a>", firstVerseLink, string.Format(Resources.Constants.VerseLinkTemplate, 1));
                    XElement verseLinksElement = new XElement(_nms + "OE",
                                                    new XElement(_nms + "T",
                                                        new XCData(StringUtils.MultiplyString("&nbsp;", 8) +
                                                            string.Join(Resources.Constants.VerseLinksDelimiter, new string[] { 
                                                                firstVerseLink + GetExistingMultiVerseString(suchNoteLink), 
                                                                OneNoteUtils.GenerateHref(ref oneNoteApp, 
                                                                    string.Format(Resources.Constants.VerseLinkTemplate, 2), notePageInfo.Id, notePageContentObjectId)
                                                                    + GetMultiVerseString(vp.ParentVersePointer ?? vp) })
                                                            )));

                    suchNoteLink.Parent.AddAfterSelf(verseLinksElement);
                }

                suchNoteLink.Value = pageLink;

                if (suchNoteLink.Parent.XPathSelectElement("one:List", notesPageDocument.Xnm) == null)  // почему то нет номера у строки
                    suchNoteLink.Parent.AddFirst(new XElement(_nms + "List",
                                                    new XElement(_nms + "Number", new XAttribute("numberSequence", 0), new XAttribute("numberFormat", "##."))));
            }

            OneNoteUtils.UpdatePageContentSafe(ref oneNoteApp, notesPageDocument.Content, notesPageDocument.Xnm);                                  

            notesPageDocument.WasModified = true;
        }

        private void RemoveExistingNoteLink(XElement suchNoteLink, OneNoteProxy.PageContent notesPageDocument)
        {
            var verseLinks = suchNoteLink.Parent.NextNode;
            if (verseLinks != null && verseLinks.XPathSelectElement("one:List", notesPageDocument.Xnm) == null)
                verseLinks.Remove();                // удаляем "ссылка1; ссылка2"

            suchNoteLink.Parent.Remove();
            suchNoteLink = null;
        }

        private XElement SearchExistingNoteLink(OneNoteProxy.PageContent notesPageDocument, HierarchyElementInfo notePageInfo, string notePageLink)
        {
            var suchNoteLink = SearchExistingNoteLink(notesPageDocument, notePageLink, _parentElement);                                        

            if (suchNoteLink == null)
            {
                //ищем в других местах
                suchNoteLink = SearchExistingNoteLink(notesPageDocument, notePageLink, null);                                        

                if (suchNoteLink != null)  // нашли в другом месте. Переносим
                {
                    var suchNoteLinkOE = suchNoteLink.Parent;
                    var suchNodeLinkSubLinks = suchNoteLinkOE.NextNode;

                    var suchNoteLinkOEChildren = suchNoteLinkOE.Parent;

                    suchNoteLinkOE.Remove();
                    _parentElement.Add(suchNoteLinkOE);  //todo: sort

                    if (suchNodeLinkSubLinks != null)
                    {
                        suchNodeLinkSubLinks.Remove();
                        _parentElement.Add(suchNodeLinkSubLinks);
                    }

                    TryDeleteTreeStructure(notesPageDocument, suchNoteLinkOEChildren); // если перенесли последнюю страницу в родителе, рекурсивно смотрим: не надо ли удалять родителей, если они стали пустыми

                    suchNoteLink = SearchExistingNoteLink(notesPageDocument, notePageLink, _parentElement);
                }
            }            

            return suchNoteLink;
        }

        private XElement SearchExistingNoteLink(OneNoteProxy.PageContent notesPageDocument, string notePageLink, XElement parentEl)
        {
            XElement suchNoteLink = null;
            string pageId;
            int pageIdStringIndex = notePageLink.IndexOf("page-id={");
            if (pageIdStringIndex == -1)
                pageIdStringIndex = notePageLink.IndexOf("{");


            var searchInAllPageString = string.Empty;
            if (parentEl == null)
            {
                searchInAllPageString = "//";
                parentEl = notesPageDocument.Content.Root;
            }

            if (pageIdStringIndex != -1)
            {
                pageId = notePageLink.Substring(pageIdStringIndex, notePageLink.IndexOf('}', pageIdStringIndex) - pageIdStringIndex + 1);
                suchNoteLink = parentEl.XPathSelectElement(string.Format("{0}one:OE/one:T[contains(.,'{1}')]", searchInAllPageString, pageId), notesPageDocument.Xnm);

                if (suchNoteLink == null)
                {
                    pageId = Uri.EscapeDataString(pageId);
                    suchNoteLink = parentEl.XPathSelectElement(
                                        string.Format("{0}one:OE/one:T[contains(translate(.,'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'),'{1}')]",
                                                    searchInAllPageString, pageId.ToUpper()), 
                                        notesPageDocument.Xnm);
                }
            }

            return suchNoteLink;
        }

        private void CreateParentTreeStructure(HierarchyElementInfo hierarchyElementInfo, XmlNamespaceManager xnm)
        {
            if (hierarchyElementInfo.Parent != null)
                CreateParentTreeStructure(hierarchyElementInfo.Parent, xnm);

            var node = _parentElement.XPathSelectElement(
                                    string.Format("one:OE/one:Meta[@name='{0}' and @content='{1}']", Consts.Constants.Key_Id, hierarchyElementInfo.Id), 
                                    xnm);

            if (node == null)
            {
                node = new XElement(_nms + "OE",
                                            new XElement(_nms + "List",
                                                        new XElement(_nms + "Number", new XAttribute("numberSequence", 0), new XAttribute("numberFormat", "##."))),
                                            new XElement(_nms + "T",
                                                new XCData(
                                                    hierarchyElementInfo.Name))
                                    );

                var childNode = new XElement(_nms + "OEChildren");
                node.Add(childNode);

                OneNoteUtils.UpdateElementMetaData(node, Constants.Key_Id, hierarchyElementInfo.Id, xnm);

                _parentElement.Add(node); //todo: sort
                _parentElement = childNode;
            }
            else
            {
                node.Parent.XPathSelectElement("one:T", xnm).Value = hierarchyElementInfo.Name;
                _parentElement = node.Parent.XPathSelectElement("one:OEChildren", xnm);
                if (_parentElement == null)  // на всякий пожарный
                {
                    var childNode = new XElement(_nms + "OEChildren");
                    node.Parent.Add(childNode);
                    _parentElement = childNode;
                }
            }
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
            var multiVerseString = string.Empty;
            var suchNoteLinkText = string.Empty;

            if (suchNoteLink != null)
                suchNoteLinkText = StringUtils.GetText(suchNoteLink.Value);

            if (!string.IsNullOrEmpty(suchNoteLinkText))
                multiVerseString = Regex.Match(suchNoteLinkText, @"\((\d+)?:\d+\-(\d+:)?\d+\)").Value;

            if (!string.IsNullOrEmpty(multiVerseString))
                return string.Format(" <b>{0}</b>", multiVerseString);

            return multiVerseString;
        }       

        private static XElement GetNotesRow(XElement tableElement, VerseNumber? verseNumber, bool isChapter, XmlNamespaceManager xnm)
        {
            XElement result = !isChapter ?
                                tableElement
                                   .XPathSelectElement(string.Format("one:Row/one:Cell[1]/one:OEChildren/one:OE/one:T[contains(.,'>:{0}<')]", verseNumber), xnm)
                              : tableElement
                                   .XPathSelectElement("one:Row/one:Cell[1]/one:OEChildren/one:OE/one:T[normalize-space(.)='']", xnm)
                                ;

            if (result != null)
                result = result.Parent.Parent.Parent.Parent;

            return result;
        }

        private static void AddNewNotesRow(ref Application oneNoteApp, VersePointer vp, bool isChapter, HierarchySearchManager.HierarchyObjectInfo verseHierarchyObjectInfo,
            XElement tableElement, XmlNamespaceManager xnm, XNamespace nms)
        {
            XElement newRow = new XElement(nms + "Row",
                                    new XElement(nms + "Cell",
                                        new XElement(nms + "OEChildren",
                                            new XElement(nms + "OE",
                                                new XElement(nms + "T",
                                                    new XCData(
                                                        !isChapter ?
                                                            OneNoteUtils.GetOrGenerateHref(ref oneNoteApp, string.Format(":{0}", verseHierarchyObjectInfo.VerseNumber),
                                                                verseHierarchyObjectInfo.VerseInfo.ObjectHref,
                                                                verseHierarchyObjectInfo.PageId, verseHierarchyObjectInfo.VerseContentObjectId,
                                                                Consts.Constants.QueryParameter_BibleVerse)
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
