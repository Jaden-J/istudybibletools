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
        internal class ListNumberInfo
        {            
            internal string NumberFormat { get; set; }
            internal int NumberSequence { get; set; }
        }

        public static readonly string Const_ManagerName = "NotesPageManagerEx";

        private XNamespace _nms;
        private HashSet<string> _processedNodes = new HashSet<string>();  // список актуализированных узлов в рамках текущей сессии анализа заметок

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
           string notesPageName, int notesPageWidth, bool force, bool processAsExtendedVerse, bool commonNotesPage, out bool rowWasAdded)
        {
            string targetContentObjectId = string.Empty;            
            
            OneNoteProxy.PageContent notesPageDocument = OneNoteProxy.Instance.GetPageContent(ref oneNoteApp, notesPageId, OneNoteProxy.PageType.NotesPage);            

            var rootElement = GetRootElementAndCreateIfNotExists(ref oneNoteApp, vp, isChapter, notesPageWidth, verseHierarchyObjectInfo,
                notesPageDocument, commonNotesPage, out rowWasAdded);

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
            OneNoteProxy.PageContent notesPageDocument, bool commonNotesPage, out bool rowWasAdded)
        {
            rowWasAdded = false;

            XElement rootElement = notesPageDocument.Content.XPathSelectElement("//one:Outline/one:OEChildren", notesPageDocument.Xnm);
            if (rootElement == null)
            {
                XElement rootElParent;
                if (commonNotesPage)
                {
                    rootElParent = new XElement(_nms + "Outline",
                                    new XElement(_nms + "OEChildren",
                                        new XElement(_nms + "OE",
                                            new XElement(_nms + "T",
                                                new XCData(
                                                     !isChapter ?
                                                        OneNoteUtils.GetOrGenerateHref(ref oneNoteApp, string.Format(":{0}", verseHierarchyObjectInfo.VerseNumber),
                                                            verseHierarchyObjectInfo.VerseInfo.ObjectHref,
                                                            verseHierarchyObjectInfo.PageId, verseHierarchyObjectInfo.VerseContentObjectId,
                                                            Consts.Constants.QueryParameter_BibleVerse)
                                                        :
                                                        string.Empty
                                                        )),
                                            new XElement(_nms + "OEChildren")
                                    )));
                    rootElement = rootElParent.XPathSelectElement("one:OEChildren/one:OE/one:OEChildren", notesPageDocument.Xnm);
                }
                else
                {
                    rootElParent = new XElement(_nms + "Outline",
                                    new XElement(_nms + "OEChildren", new XAttribute("indent", 2)));
                    rootElement = rootElParent.XPathSelectElement("one:OEChildren", notesPageDocument.Xnm);
                }

                notesPageDocument.Content.Root.Add(rootElParent);                               
            }

            return rootElement;
        }

        public string GetNotesRowObjectId(ref Application oneNoteApp, string notesPageId, VerseNumber? verseNumber, bool isChapter)
        {
            var result = string.Empty;
            var notesPageDocument = OneNoteProxy.Instance.GetPageContent(ref oneNoteApp, notesPageId, OneNoteProxy.PageType.NotesPage);
            var targetElement = notesPageDocument.Content.Root.XPathSelectElement("one:Title/one:OE", notesPageDocument.Xnm);

            if (targetElement != null)
                result = (string)targetElement.Attribute("objectID");

            return result;
        }

        private XElement _parentElement;
        private int _level;

        private void AddLinkToNotesPage(ref Application oneNoteApp, NoteLinkManager noteLinkManager, VersePointer vp, XElement rootElement,
           HierarchyElementInfo notePageInfo, string notePageContentObjectId,
           OneNoteProxy.PageContent notesPageDocument, string notesPageName, bool force, bool processAsExtendedVerse)
        {
            _parentElement = rootElement;
            _level = 1;

            if (notePageInfo.Parent != null)
                CreateParentTreeStructure(ref oneNoteApp, notePageInfo.Parent, notesPageDocument.Xnm);

            string link = OneNoteUtils.GenerateHref(ref oneNoteApp, notePageInfo.Name, notePageInfo.Id, notePageContentObjectId);

            var suchNoteLink = SearchExistingNoteLink(ref oneNoteApp, notesPageDocument, notePageInfo, link);

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
                var listNumberInfo = GetListNumberInfo(_level);
                XElement linkElement = new XElement(_nms + "OE",
                                            new XElement(_nms + "List",
                                                        new XElement(_nms + "Number",
                                                            new XAttribute("numberSequence", listNumberInfo.NumberSequence),
                                                            new XAttribute("numberFormat", listNumberInfo.NumberFormat))),
                                            new XElement(_nms + "T",
                                                new XCData(
                                                    link + GetMultiVerseString(vp.ParentVersePointer ?? vp))));
                OneNoteUtils.UpdateElementMetaData(linkElement, Constants.Key_Id, notePageInfo.Id, notesPageDocument.Xnm);

                TryToInsertOrMoveElement(ref oneNoteApp, linkElement, notePageInfo, _parentElement, MoveOperationType.Insert, notesPageDocument.Xnm);
            }
            else if (!processAsExtendedVerse)
            {   
                if (!_processedNodes.Contains(notePageInfo.Id))
                {
                    suchNoteLink = TryToInsertOrMoveElement(ref oneNoteApp, suchNoteLink.Parent, notePageInfo, _parentElement, MoveOperationType.MoveWithLinksRow, notesPageDocument.Xnm)
                                            .XPathSelectElement("one:T", notesPageDocument.Xnm);
                    _processedNodes.Add(notePageInfo.Id);
                }

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
                {
                    var listNumberInfo = GetListNumberInfo(_level);
                    suchNoteLink.Parent.AddFirst(new XElement(_nms + "List",
                                                    new XElement(_nms + "Number",
                                                            new XAttribute("numberSequence", listNumberInfo.NumberSequence),
                                                            new XAttribute("numberFormat", listNumberInfo.NumberFormat))));
                }
            }

            //OneNoteUtils.UpdatePageContentSafe(ref oneNoteApp, notesPageDocument.Content, notesPageDocument.Xnm);                                  

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

        private XElement SearchExistingNoteLink(ref Application oneNoteApp, OneNoteProxy.PageContent notesPageDocument, HierarchyElementInfo notePageInfo, string notePageLink)
        {
            var suchNoteLink = SearchExistingNoteLinkInParent(notesPageDocument, notePageLink, _parentElement);                                        

            if (suchNoteLink == null)
            {
                //ищем в других местах
                suchNoteLink = SearchExistingNoteLinkInParent(notesPageDocument, notePageLink, null);                                        

                if (suchNoteLink != null)  // нашли в другом месте. Переносим
                {
                    var suchNoteLinkOE = suchNoteLink.Parent;
                    var suchNoteLinkOEChildren = suchNoteLinkOE.Parent;

                    TryToInsertOrMoveElement(ref oneNoteApp, suchNoteLinkOE, notePageInfo, _parentElement, MoveOperationType.MoveWithLinksRow, notesPageDocument.Xnm);
                    if (!_processedNodes.Contains(notePageInfo.Id))
                        _processedNodes.Add(notePageInfo.Id);  // чтоб больше не обрабатывали

                    TryToDeleteTreeStructure(suchNoteLinkOEChildren); // если перенесли последнюю страницу в родителе, рекурсивно смотрим: не надо ли удалять родителей, если они стали пустыми

                    suchNoteLink = SearchExistingNoteLinkInParent(notesPageDocument, notePageLink, _parentElement);

                    // перенесли узел с другого уровня скорее всего. обновляем символ нумерованного списка                    
                    var number = suchNoteLink.Parent.XPathSelectElement("one:List/one:Number", notesPageDocument.Xnm);
                    if (number != null)
                    {
                        var listNumberInfo = GetListNumberInfo(_level);
                        number.SetAttributeValue("numberSequence", listNumberInfo.NumberSequence);
                        number.SetAttributeValue("numberFormat", listNumberInfo.NumberFormat);
                    }
                }
            }            

            return suchNoteLink;
        }

        private void TryToDeleteTreeStructure(XElement suchNoteLinkOEChildren)
        {
            if (suchNoteLinkOEChildren.Nodes().Count() == 0)
            {
                var grandParent = suchNoteLinkOEChildren.Parent.Parent;
                suchNoteLinkOEChildren.Parent.Remove();

                TryToDeleteTreeStructure(grandParent);
            }
        }

        private XElement SearchExistingNoteLinkInParent(OneNoteProxy.PageContent notesPageDocument, string notePageLink, XElement parentEl)
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

        private void CreateParentTreeStructure(ref Application oneNoteApp, HierarchyElementInfo hierarchyElementInfo, XmlNamespaceManager xnm)
        {
            if (hierarchyElementInfo.Parent != null)
                CreateParentTreeStructure(ref oneNoteApp, hierarchyElementInfo.Parent, xnm);

            var node = _parentElement.XPathSelectElement(
                                    string.Format("one:OE/one:Meta[@name='{0}' and @content='{1}']", Consts.Constants.Key_Id, hierarchyElementInfo.Id), 
                                    xnm);

            if (node == null)
            {
                var listNumberInfo = GetListNumberInfo(_level);
                node = new XElement(_nms + "OE",
                                            new XElement(_nms + "List",
                                                        new XElement(_nms + "Number",
                                                            new XAttribute("numberSequence", listNumberInfo.NumberSequence),
                                                            new XAttribute("numberFormat", listNumberInfo.NumberFormat))),
                                            new XElement(_nms + "T",
                                                new XCData(
                                                    hierarchyElementInfo.Name))
                                    );

                var childNode = new XElement(_nms + "OEChildren");
                node.Add(childNode);

                OneNoteUtils.UpdateElementMetaData(node, Constants.Key_Id, hierarchyElementInfo.Id, xnm);

                TryToInsertOrMoveElement(ref oneNoteApp, node, hierarchyElementInfo, _parentElement, MoveOperationType.Insert, xnm);

                _parentElement = childNode;                
            }
            else
            {
                node = node.Parent;
                if (!_processedNodes.Contains(hierarchyElementInfo.Id))
                {
                    node.XPathSelectElement("one:T", xnm).Value = hierarchyElementInfo.Name;                    

                    TryToInsertOrMoveElement(ref oneNoteApp, node, hierarchyElementInfo, _parentElement, MoveOperationType.Move, xnm);

                    _processedNodes.Add(hierarchyElementInfo.Id);
                }
                
                _parentElement = node.XPathSelectElement("one:OEChildren", xnm);                
                if (_parentElement == null)  // на всякий пожарный
                {
                    var childNode = new XElement(_nms + "OEChildren");
                    node.Add(childNode);
                    _parentElement = childNode;
                }
            }
            _level++;
        }

  
        private enum MoveOperationType
        {
            Insert,
            Move,
            MoveWithLinksRow
        }
        private static XElement TryToInsertOrMoveElement(ref Application oneNoteApp, XElement el, HierarchyElementInfo elInfo, 
                                                        XElement parentEl, MoveOperationType moveType, XmlNamespaceManager xnm)
        {
            bool linkWasFound;
            var prevLink = GetPrevNoteLink(ref oneNoteApp, elInfo, parentEl, xnm, out linkWasFound);

            var needToMoveOrInsert = !(linkWasFound && prevLink == null);  // иначе ссылка стоит в начале и она должана там стоять
            if (needToMoveOrInsert && linkWasFound)
            {
                var realPrevLink = GetRealPrevLink(el, xnm);
                if (realPrevLink != null)
                {
                    if (realPrevLink == prevLink)  // ссылка и так уже не правильном месте
                        needToMoveOrInsert = false;
                }
            }

            if (needToMoveOrInsert)  // иначе ссылка стоит на правильном месте
            {   
                XNode linksRow = null;
                if (moveType == MoveOperationType.Move || moveType == MoveOperationType.MoveWithLinksRow)  
                {
                    if (moveType == MoveOperationType.MoveWithLinksRow)
                    {
                        if (el.NextNode != null && el.NextNode.XPathSelectElement("one:List", xnm) == null)  // если следующая строка типа "ссылка1; ссылка2"                                    
                            linksRow = el.NextNode;
                    }

                    el.Remove();

                    if (linksRow != null)
                        linksRow.Remove();
                }

                el = InsertElement(el, elInfo, parentEl, prevLink, linksRow, moveType == MoveOperationType.MoveWithLinksRow, xnm);                
            }

            return el;
        }

        private static XNode GetRealPrevLink(XElement el, XmlNamespaceManager xnm)
        {
            var result = el.PreviousNode;

            if (result != null)
                if (result.XPathSelectElement("one:List", xnm) == null)
                    result = result.PreviousNode;

            return result;
        }

        private static XNode GetPrevNoteLink(ref Application oneNoteApp, HierarchyElementInfo elInfo, XElement parentEl, XmlNamespaceManager xnm, out bool linkWasFound)
        {
            XElement prevLink = null;
            linkWasFound = false;

            var notebookHierarchy = OneNoteProxy.Instance.GetHierarchy(ref oneNoteApp, elInfo.NotebookId, HierarchyScope.hsPages);  //from cache
            XElement parentHierarchy;
            if (elInfo.Type != HierarchyElementType.Notebook)
                parentHierarchy = notebookHierarchy.Content.Root.XPathSelectElement(string.Format("//one:{0}[@ID='{1}']", elInfo.GetElementName(), elInfo.Id), xnm).Parent;
            else
                parentHierarchy = OneNoteProxy.Instance.GetHierarchy(ref oneNoteApp, null, HierarchyScope.hsNotebooks).Content.Root;

            var noteLinkInHierarchy = parentHierarchy.XPathSelectElement(string.Format("*[@ID='{0}']", elInfo.Id), xnm);

            var prevNodesInHierarchy = noteLinkInHierarchy.NodesBeforeSelf();

            if (prevNodesInHierarchy.Count() != 0)
            {
                foreach (XElement existingLink in parentEl.XPathSelectElements("one:OE/one:List", xnm))
                {                    
                    var existingLinkId = OneNoteUtils.GetElementMetaData(existingLink.Parent, Constants.Key_Id, xnm);

                    if (existingLinkId == elInfo.Id)
                    {
                        linkWasFound = true;
                    }
                    else
                    {
                        var existingLinkInHierarchy = parentHierarchy.XPathSelectElement(string.Format("*[@ID='{0}']", existingLinkId), xnm);
                        if (!prevNodesInHierarchy.Contains(existingLinkInHierarchy))
                            break;

                        prevLink = existingLink.Parent;
                    }
                }
            }

            return prevLink;
        }

        private static XElement InsertElement(XElement el, HierarchyElementInfo elInfo, XElement parentElement, XNode prevLink, XNode linksRow, bool updateMetadata, XmlNamespaceManager xnm)
        { 
            if (prevLink == null)
            {
                if (linksRow != null)
                    parentElement.AddFirst(linksRow);               

                parentElement.AddFirst(el);                        
            }
            else
            {
                if (prevLink.NextNode != null && prevLink.NextNode.XPathSelectElement("one:List", xnm) == null)  // если следующая строка типа "ссылка1; ссылка2;"                    
                    prevLink = prevLink.NextNode;

                if (linksRow != null)
                    prevLink.AddAfterSelf(linksRow);

                prevLink.AddAfterSelf(el);
            }

            if (updateMetadata)
                OneNoteUtils.UpdateElementMetaData(el, Constants.Key_Id, elInfo.Id, xnm);

            return el;
        }



        private static ListNumberInfo GetListNumberInfo(int level)
        {
            var result = new ListNumberInfo();
            
            switch (level)
            {
                case 1:
                    //result.Text = "I.";
                    result.NumberFormat = "##.";
                    result.NumberSequence = 1;
                    break;
                case 2:
                    //result.Text = "A.";
                    result.NumberFormat = "##.";
                    result.NumberSequence = 3;
                    break;
                case 3:
                    //result.Text = "1.";
                    result.NumberFormat = "##.";
                    result.NumberSequence = 0;
                    break;
                case 4:
                    //result.Text = "a.";
                    result.NumberFormat = "##.";
                    result.NumberSequence = 4;
                    break;
                case 5:
                    //result.Text = "1)";
                    result.NumberFormat = "##)";
                    result.NumberSequence = 0;
                    break;
                case 6:
                    //result.Text = "a)";
                    result.NumberFormat = "##)";
                    result.NumberSequence = 4;
                    break;
                case 7:
                    //result.Text = "(1)";
                    result.NumberFormat = "(##)";
                    result.NumberSequence = 0;
                    break;
                case 8:
                    //result.Text = "(a)";
                    result.NumberFormat = "(##)";
                    result.NumberSequence = 4;
                    break;
                case 9:
                    //result.Text = "1 &gt;";
                    result.NumberFormat = "## &gt;";
                    result.NumberSequence = 0;
                    break;
                case 10:
                    //result.Text = "a &gt;";
                    result.NumberFormat = "## &gt;";
                    result.NumberSequence = 4;
                    break;
                default:
                    //result.Text = "1.";
                    result.NumberFormat = "##.";
                    result.NumberSequence = 0;
                    break;
            }

            return result;
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
    }
}
