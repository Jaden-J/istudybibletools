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
using System.Text.RegularExpressions;
using BibleCommon.Contracts;

namespace BibleCommon.Services
{
    public class NotesPageManagerEx : INotesPageManager
    {
        private enum MoveOperationType
        {
            Insert,
            Move,
            MoveAndUpdateMetadata
        }

        internal class ListNumberInfo
        {
            internal string NumberFormat { get; set; }
            internal int NumberSequence { get; set; }
        }

        public static readonly string Const_ManagerName = "NotesPageManagerEx";

        private XNamespace _nms;
        private HashSet<string> _processedNodes = new HashSet<string>();  // список актуализированных узлов в рамках текущей сессии анализа заметок
        private Dictionary<string, int?> _notebooksDisplayLevel = new Dictionary<string, int?>();

        public string ManagerName
        {
            get { return Const_ManagerName; }
        }

        public NotesPageManagerEx()
        {
            _nms = XNamespace.Get(Constants.OneNoteXmlNs);
            foreach (var notebookInfo in SettingsManager.Instance.SelectedNotebooksForAnalyze)
                _notebooksDisplayLevel.Add(notebookInfo.NotebookId, notebookInfo.DisplayLevels);
        }

        public string UpdateNotesPage(ref Application oneNoteApp, NoteLinkManager noteLinkManager, VersePointer vp, int versePosition,
           bool isChapter, HierarchySearchManager.HierarchyObjectInfo verseHierarchyObjectInfo,
           HierarchyElementInfo notePageInfo, string notesPageId, string notePageContentObjectId,
           string notesPageName, int notesPageWidth, bool force, bool processAsExtendedVerse, out bool rowWasAdded)
        {
            string targetContentObjectId = string.Empty;

            OneNoteProxy.PageContent notesPageDocument = OneNoteProxy.Instance.GetPageContent(ref oneNoteApp, notesPageId, OneNoteProxy.PageType.NotesPage);            

            var rootElement = GetVerseRootElementAndCreateIfNotExists(ref oneNoteApp, vp, isChapter, notesPageWidth, verseHierarchyObjectInfo,
                notesPageDocument, out rowWasAdded);

            if (rootElement != null)
            {
                AddLinkToNotesPage(ref oneNoteApp, noteLinkManager, vp, versePosition, rootElement, notePageInfo,
                    notePageContentObjectId, notesPageDocument, notesPageName, notePageInfo.NotebookId, force, processAsExtendedVerse);

                targetContentObjectId = GetNotesRowObjectId(notesPageDocument, notesPageId, verseHierarchyObjectInfo.VerseNumber);
            }

            return targetContentObjectId;
        }

        public string GetNotesRowObjectId(ref Application oneNoteApp, string notesPageId, VerseNumber? verseNumber, bool isChapter)
        {
            var notesPageDocument = OneNoteProxy.Instance.GetPageContent(ref oneNoteApp, notesPageId, OneNoteProxy.PageType.NotesPage);

            return GetNotesRowObjectId(notesPageDocument, notesPageId, verseNumber);
        }


        private XElement GetVerseRootElementAndCreateIfNotExists(ref Application oneNoteApp, VersePointer vp, bool isChapter,
           int mainColumnWidth, HierarchySearchManager.HierarchyObjectInfo verseHierarchyObjectInfo,
           OneNoteProxy.PageContent notesPageDocument, out bool rowWasAdded)
        {
            rowWasAdded = false;

            var rootElement = notesPageDocument.Content.Root.XPathSelectElement(
                                            string.Format("one:Outline/one:OEChildren/one:OE/one:Meta[@name='{0}' and @content='{1}']", Constants.Key_Verse, vp.Verse.GetValueOrDefault(0)),
                                            notesPageDocument.Xnm);


            if (rootElement == null)
            {
                var sizeEl = new XElement(_nms + "Size", new XAttribute("width", mainColumnWidth), new XAttribute("height", 15), new XAttribute("isSetByUser", true));

                rootElement = CreateRootElelementForNotesPage(ref oneNoteApp, vp, isChapter, sizeEl, verseHierarchyObjectInfo, notesPageDocument);

                rowWasAdded = true;
            }
            else
            {
                rootElement = rootElement.Parent.XPathSelectElement("one:OEChildren", notesPageDocument.Xnm);
            }

            return rootElement;
        }      

        private XElement CreateRootElelementForNotesPage(ref Application oneNoteApp, VersePointer vp, bool isChapter, XElement sizeEl,
            HierarchySearchManager.HierarchyObjectInfo verseHierarchyObjectInfo, OneNoteProxy.PageContent notesPageDocument)
        {
            var rootElParent = notesPageDocument.Content.Root.XPathSelectElement("one:Outline/one:OEChildren", notesPageDocument.Xnm);

            if (rootElParent == null)
            {
                rootElParent = notesPageDocument.Content.Root.XPathSelectElement("one:Outline", notesPageDocument.Xnm);

                if (rootElParent == null)
                {
                    rootElParent = new XElement(_nms + "Outline", sizeEl);
                    notesPageDocument.Content.Root.Add(rootElParent);
                }

                rootElParent.Add(new XElement(_nms + "OEChildren"));                

                rootElParent = rootElParent.XPathSelectElement("one:OEChildren", notesPageDocument.Xnm);
            }

            var rootElement = new XElement(_nms + "OE",
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
                            );
            OneNoteUtils.UpdateElementMetaData(rootElement, Constants.Key_Verse, vp.Verse.GetValueOrDefault(0).ToString(), notesPageDocument.Xnm);

            XElement prevOE = null;
            if (!isChapter)  // иначе добавляем первым
            {
                foreach (var oeMEta in rootElParent.XPathSelectElements(string.Format("one:OE/one:Meta[@name='{0}']", Constants.Key_Verse), notesPageDocument.Xnm))
                {
                    var verse = int.Parse((string)oeMEta.Attribute("content"));
                    if (verse > vp.Verse)
                        break;

                    prevOE = oeMEta.Parent;
                }
            }

            if (prevOE == null)
                rootElParent.AddFirst(rootElement);
            else
                prevOE.AddAfterSelf(rootElement);            

            return rootElement.XPathSelectElement("one:OEChildren", notesPageDocument.Xnm);
        }

        private string GetNotesRowObjectId(OneNoteProxy.PageContent notesPageDocument, string notesPageId, VerseNumber? verseNumber)
        {
            var result = string.Empty;
            XElement targetElement = null;

            targetElement = notesPageDocument.Content.Root.XPathSelectElement(
                                        string.Format("one:Outline/one:OEChildren/one:OE/one:Meta[@name='{0}' and @content='{1}']",
                                                                Constants.Key_Verse, verseNumber.GetValueOrDefault(new VerseNumber(0)).Verse),
                                        notesPageDocument.Xnm);

            if (targetElement != null)
                targetElement = targetElement.Parent;

            if (targetElement == null)
                targetElement = notesPageDocument.Content.Root.XPathSelectElement("one:Title/one:OE", notesPageDocument.Xnm);

            if (targetElement != null)
                result = (string)targetElement.Attribute("objectID");

            return result;
        }     

        private XElement _parentElement;
        private int _level;

        private void AddLinkToNotesPage(ref Application oneNoteApp, NoteLinkManager noteLinkManager, VersePointer vp, int versePosition,
           XElement rootElement, HierarchyElementInfo notePageInfo, string notePageContentObjectId,
           OneNoteProxy.PageContent notesPageDocument, string notesPageName, string notebookId, bool force, bool processAsExtendedVerse)
        {
            _parentElement = rootElement;
            _level = 1;

            if (notePageInfo.Parent != null)
                CreateParentTreeStructure(ref oneNoteApp, notePageInfo.Parent, notebookId, notesPageDocument.Xnm);

            string link = OneNoteUtils.GenerateHref(ref oneNoteApp, notePageInfo.Name, notePageInfo.Id, notePageContentObjectId,
                string.Format("{0}={1}", Constants.QueryParameterKey_VersePosition, versePosition));

            var suchNoteLink = SearchExistingNoteLink(ref oneNoteApp, rootElement, notePageInfo, link, notesPageDocument.Xnm);

            if (suchNoteLink != null)
            {
                var key = new NoteLinkManager.NotePageProcessedVerseId() { NotePageId = notePageInfo.Id, NotesPageName = notesPageName };
                if (force && !noteLinkManager.ContainsNotePageProcessedVerse(key, vp) && !processAsExtendedVerse)  // если в первый раз и force и не расширенный стих
                {  // удаляем старые ссылки на текущую странцу, так как мы начали новый анализ с параметром "force" и мы только в первый раз зашли сюда
                    suchNoteLink.Parent.Remove();
                    suchNoteLink = null;
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
                    suchNoteLink = TryToInsertOrMoveElement(ref oneNoteApp, suchNoteLink.Parent, notePageInfo, _parentElement, MoveOperationType.MoveAndUpdateMetadata, notesPageDocument.Xnm)
                                            .XPathSelectElement("one:T", notesPageDocument.Xnm);
                    _processedNodes.Add(notePageInfo.Id);
                }

                string pageLink = OneNoteUtils.GenerateHref(ref oneNoteApp, notePageInfo.Name, notePageInfo.Id, notePageInfo.PageTitleId);

                var existingVerseLinksElement = suchNoteLink.Parent.XPathSelectElement("one:OEChildren/one:OE/one:T", notesPageDocument.Xnm);
                if (existingVerseLinksElement != null)
                {
                    InsertAdditionalVerseLink(ref oneNoteApp, existingVerseLinksElement, notePageInfo, notePageContentObjectId, vp, versePosition);
                }
                else  // значит мы нашли второе упоминание данной ссылки в заметке
                {
                    InsertSecondVerseLink(ref oneNoteApp, suchNoteLink, notePageInfo, notePageContentObjectId, vp, versePosition);                    
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

        private void InsertSecondVerseLink(ref Application oneNoteApp, XElement suchNoteLink, HierarchyElementInfo notePageInfo,
            string notePageContentObjectId, VersePointer vp, int versePosition)
        {
            var firstVerseLink = StringUtils.GetAttributeValue(suchNoteLink.Value, "href");
            var firstVersePosition = GetVersePosition(suchNoteLink.Value);
            

            firstVerseLink = string.Format("<a href='{0}'>{1}</a>", firstVerseLink, string.Format(Resources.Constants.VerseLinkTemplate,
                firstVersePosition > versePosition ? 2 : 1)) + GetExistingMultiVerseString(suchNoteLink);

            var verseLink = OneNoteUtils.GenerateHref(ref oneNoteApp,
                                            string.Format(Resources.Constants.VerseLinkTemplate, firstVersePosition > versePosition ? 1 : 2), 
                                            notePageInfo.Id, notePageContentObjectId,
                                            string.Format("{0}={1}", Constants.QueryParameterKey_VersePosition, versePosition))
                                        + GetMultiVerseString(vp.ParentVersePointer ?? vp);

            var arrayOfLinks = firstVersePosition > versePosition
                                    ? new string[] { verseLink, firstVerseLink }
                                    : new string[] { firstVerseLink, verseLink };


            var verseLinksElement = new XElement(_nms + "OEChildren",
                                            new XElement(_nms + "OE",
                                                new XElement(_nms + "T",
                                                    new XCData(
                                                        string.Join(Resources.Constants.VerseLinksDelimiter, arrayOfLinks
                                                        )))));

            suchNoteLink.Parent.Add(verseLinksElement);
        }

        private static int? GetVersePosition(string href)
        {
            var s = StringUtils.GetQueryParameterValue(href, Constants.QueryParameterKey_VersePosition);
            if (!string.IsNullOrEmpty(s))
                return int.Parse(s);

            return null;
        }

        private void InsertAdditionalVerseLink(ref Application oneNoteApp, XElement existingVerseLinksElement, HierarchyElementInfo notePageInfo, 
            string notePageContentObjectId, VersePointer vp, int versePointeHtmlStartIndex)
        {
            int currentVerseIndex = existingVerseLinksElement.Value.Split(new string[] { "</a>" }, StringSplitOptions.None).Length;

            existingVerseLinksElement.Value += Resources.Constants.VerseLinksDelimiter + OneNoteUtils.GenerateHref(ref oneNoteApp,
                        string.Format(Resources.Constants.VerseLinkTemplate, currentVerseIndex), notePageInfo.Id, notePageContentObjectId)
                        + GetMultiVerseString(vp.ParentVersePointer ?? vp);
        }

        private XElement SearchExistingNoteLink(ref Application oneNoteApp, XElement rootElement, HierarchyElementInfo notePageInfo, string notePageLink, XmlNamespaceManager xnm)
        {
            var suchNoteLink = SearchExistingNoteLinkInParent(_parentElement, rootElement, notePageLink, xnm);

            if (suchNoteLink == null)
            {
                //ищем в других местах
                suchNoteLink = SearchExistingNoteLinkInParent(null, rootElement, notePageLink, xnm);

                if (suchNoteLink != null)  // нашли в другом месте. Переносим
                {
                    var suchNoteLinkOE = suchNoteLink.Parent;
                    var suchNoteLinkOEChildren = suchNoteLinkOE.Parent;

                    TryToInsertOrMoveElement(ref oneNoteApp, suchNoteLinkOE, notePageInfo, _parentElement, MoveOperationType.MoveAndUpdateMetadata, xnm);
                    if (!_processedNodes.Contains(notePageInfo.Id))
                        _processedNodes.Add(notePageInfo.Id);  // чтоб больше не обрабатывали

                    TryToDeleteTreeStructure(suchNoteLinkOEChildren); // если перенесли последнюю страницу в родителе, рекурсивно смотрим: не надо ли удалять родителей, если они стали пустыми

                    suchNoteLink = SearchExistingNoteLinkInParent(_parentElement, rootElement, notePageLink, xnm);

                    // перенесли узел с другого уровня скорее всего. обновляем символ нумерованного списка                    
                    var number = suchNoteLink.Parent.XPathSelectElement("one:List/one:Number", xnm);
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

        private XElement SearchExistingNoteLinkInParent(XElement parentEl, XElement rootElement, string notePageLink, XmlNamespaceManager xnm)
        {
            XElement suchNoteLink = null;
            string pageId;
            int pageIdStringIndex = notePageLink.IndexOf("page-id={");
            if (pageIdStringIndex == -1)
                pageIdStringIndex = notePageLink.IndexOf("{");


            var searchInAllPageString = string.Empty;
            if (parentEl == null)
            {
                searchInAllPageString = ".//";
                parentEl = rootElement;
            }

            if (pageIdStringIndex != -1)
            {
                pageId = notePageLink.Substring(pageIdStringIndex, notePageLink.IndexOf('}', pageIdStringIndex) - pageIdStringIndex + 1);
                suchNoteLink = parentEl.XPathSelectElement(string.Format("{0}one:OE/one:T[contains(.,'{1}')]", searchInAllPageString, pageId), xnm);

                if (suchNoteLink == null)
                {
                    pageId = Uri.EscapeDataString(pageId);
                    suchNoteLink = parentEl.XPathSelectElement(
                                        string.Format("{0}one:OE/one:T[contains(translate(.,'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'),'{1}')]",
                                                    searchInAllPageString, pageId.ToUpper()),
                                        xnm);
                }
            }

            return suchNoteLink;
        }

        private void CreateParentTreeStructure(ref Application oneNoteApp, HierarchyElementInfo hierarchyElementInfo, string notebookId, XmlNamespaceManager xnm)
        {
            if (hierarchyElementInfo.Parent != null)
                CreateParentTreeStructure(ref oneNoteApp, hierarchyElementInfo.Parent, notebookId, xnm);

            var node = _parentElement.XPathSelectElement(
                                    string.Format("one:OE/one:Meta[@name='{0}' and @content='{1}']", Consts.Constants.Key_Id, hierarchyElementInfo.Id),
                                    xnm);

            if (node == null)
            {
                var listNumberInfo = GetListNumberInfo(_level);
                int? displayLevel = null;
                if (_notebooksDisplayLevel.ContainsKey(notebookId))
                    displayLevel = _notebooksDisplayLevel[notebookId];

                node = new XElement(_nms + "OE",
                                            new XElement(_nms + "List",
                                                        new XElement(_nms + "Number",
                                                            new XAttribute("numberSequence", listNumberInfo.NumberSequence),
                                                            new XAttribute("numberFormat", listNumberInfo.NumberFormat))),
                                            new XElement(_nms + "T",
                                                new XCData(
                                                    hierarchyElementInfo.Name))
                                    );

                if (displayLevel != null && _level >= displayLevel)
                    node.Add(new XAttribute("collapsed", 1));

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

        private static XElement TryToInsertOrMoveElement(ref Application oneNoteApp, XElement el, HierarchyElementInfo elInfo,
                                                        XElement parentEl, MoveOperationType moveType, XmlNamespaceManager xnm)
        {
            bool linkWasFound;
            var prevLink = GetPrevNoteLink(ref oneNoteApp, elInfo, parentEl, xnm, out linkWasFound);

            var needToMoveOrInsert = !(linkWasFound && prevLink == null);  // иначе ссылка стоит в начале и она должана там стоять
            if (needToMoveOrInsert && linkWasFound)
            {
                if (el.PreviousNode == prevLink)  // ссылка и так уже на правильном месте
                    needToMoveOrInsert = false;
            }

            if (needToMoveOrInsert)  // иначе ссылка стоит на правильном месте
            {
                if (moveType == MoveOperationType.Move || moveType == MoveOperationType.MoveAndUpdateMetadata)
                {
                    el.Remove();
                }

                el = InsertElement(el, elInfo, parentEl, prevLink, moveType == MoveOperationType.MoveAndUpdateMetadata, xnm);
            }

            return el;
        }

        private static XNode GetRealPrevLink(XElement el, XmlNamespaceManager xnm)
        {
            var result = el.PreviousNode;


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
                foreach (var existingLink in parentEl.XPathSelectElements("one:OE", xnm))
                {
                    var existingLinkId = OneNoteUtils.GetElementMetaData(existingLink, Constants.Key_Id, xnm);

                    if (existingLinkId == elInfo.Id)
                    {
                        linkWasFound = true;
                    }
                    else
                    {
                        var existingLinkInHierarchy = parentHierarchy.XPathSelectElement(string.Format("*[@ID='{0}']", existingLinkId), xnm);
                        if (!prevNodesInHierarchy.Contains(existingLinkInHierarchy))
                            break;

                        prevLink = existingLink;
                    }
                }
            }

            return prevLink;
        }

        private static XElement InsertElement(XElement el, HierarchyElementInfo elInfo, XElement parentElement, XNode prevLink, bool updateMetadata, XmlNamespaceManager xnm)
        {
            if (prevLink == null)
                parentElement.AddFirst(el);
            else
                prevLink.AddAfterSelf(el);

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