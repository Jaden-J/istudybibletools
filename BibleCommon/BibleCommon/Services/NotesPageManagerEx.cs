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
        private Dictionary<string, HashSet<string>> _processedNodes = new Dictionary<string, HashSet<string>>();  // список актуализированных узлов в рамках текущей сессии анализа заметок
        private Dictionary<string, int?> _notebooksDisplayLevel = new Dictionary<string, int?>();

        public string ManagerName
        {
            get { return Const_ManagerName; }
        }

        public NotesPageManagerEx()
        {
            _nms = XNamespace.Get(Constants.OneNoteXmlNs);
            foreach (var notebookInfo in SettingsManager.Instance.SelectedNotebooksForAnalyze)
                if (!_notebooksDisplayLevel.ContainsKey(notebookInfo.NotebookId))  // на всякий пожарный
                    _notebooksDisplayLevel.Add(notebookInfo.NotebookId, notebookInfo.DisplayLevels);
        }

        public string UpdateNotesPage(ref Application oneNoteApp, NoteLinkManager noteLinkManager, VersePointer vp, 
           decimal verseWeight, XmlCursorPosition versePosition,
           bool isChapter, HierarchySearchManager.HierarchyObjectInfo verseHierarchyObjectInfo,
           HierarchyElementInfo notePageInfo, string notesPageId, string notePageContentObjectId,
           string notesPageName, int notesPageWidth, bool isImportantVerse, bool force, bool processAsExtendedVerse, out bool rowWasAdded)
        {
            string targetContentObjectId = string.Empty;

            OneNoteProxy.PageContent notesPageDocument = OneNoteProxy.Instance.GetPageContent(ref oneNoteApp, notesPageId, OneNoteProxy.PageType.NotesPage);            

            var rootElement = GetVerseRootElementAndCreateIfNotExists(ref oneNoteApp, vp, isChapter, notesPageWidth, verseHierarchyObjectInfo,
                notesPageDocument, out rowWasAdded);

            if (rootElement != null)
            {
                AddLinkToNotesPage(ref oneNoteApp, noteLinkManager, vp, verseWeight, versePosition, rootElement, notePageInfo,
                    notePageContentObjectId, notesPageDocument, notesPageName, notePageInfo.NotebookId, isImportantVerse, force, processAsExtendedVerse);

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
                                            string.Format("one:Outline/one:OEChildren/one:OE/one:Meta[@name=\"{0}\" and @content=\"{1}\"]", Constants.Key_Verse, vp.Verse.GetValueOrDefault(0)),
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
                                            OneNoteUtils.GetOrGenerateLink(ref oneNoteApp, string.Format(":{0}", verseHierarchyObjectInfo.VerseNumber),
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
                foreach (var oeMEta in rootElParent.XPathSelectElements(string.Format("one:OE/one:Meta[@name=\"{0}\"]", Constants.Key_Verse), notesPageDocument.Xnm))
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
                                        string.Format("one:Outline/one:OEChildren/one:OE/one:Meta[@name=\"{0}\" and @content=\"{1}\"]",
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

        private void AddLinkToNotesPage(ref Application oneNoteApp, NoteLinkManager noteLinkManager, VersePointer vp, decimal verseWeight, XmlCursorPosition versePosition,
           XElement rootElement, HierarchyElementInfo notePageInfo, string notePageContentObjectId,
           OneNoteProxy.PageContent notesPageDocument, string notesPageName, string notebookId, bool isImportantVerse, bool force, bool processAsExtendedVerse)
        {
            _parentElement = rootElement;
            _level = 1;

            if (notePageInfo.Parent != null)
                CreateParentTreeStructure(ref oneNoteApp, notePageInfo.Parent, notebookId, notesPageName, notesPageDocument.Xnm);

            var linkArgs = new List<string>();
            linkArgs.Add(string.Format("{0}={1}", Constants.QueryParameterKey_VersePosition, versePosition));
            linkArgs.Add(string.Format("{0}={1}", Constants.QueryParameterKey_VerseWeight, verseWeight));
            if (!string.IsNullOrEmpty(notePageInfo.UniqueId))
                linkArgs.Add(string.Format("{0}={1}", Constants.QueryParameterKey_NotePageId, notePageInfo.UniqueId));

            string link = OneNoteUtils.GenerateLink(ref oneNoteApp, 
                            GetVerseLinkTitle(notePageInfo.UniqueTitle, verseWeight >= Constants.ImportantVerseWeight), 
                            notePageInfo.Id, notePageContentObjectId, linkArgs.ToArray());

            var suchNoteLink = SearchExistingNoteLink(ref oneNoteApp, rootElement, notePageInfo, link, notesPageName, notesPageDocument.Xnm);

            if (suchNoteLink != null)
            {
                var key = new NoteLinkManager.NotePageProcessedVerseId() { NotePageId = notePageInfo.UniqueId ?? notePageInfo.Id, NotesPageName = notesPageName };
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
                                                            new XAttribute("bold", verseWeight >= Constants.ImportantVerseWeight),
                                                            new XAttribute("numberFormat", listNumberInfo.NumberFormat))),
                                            new XElement(_nms + "T",
                                                new XCData(
                                                    link + GetMultiVerseString(vp.ParentVersePointer ?? vp))));
                OneNoteUtils.UpdateElementMetaData(linkElement, Constants.Key_Id, notePageInfo.UniqueName, notesPageDocument.Xnm);

                TryToInsertOrMoveElement(ref oneNoteApp, linkElement, notePageInfo, _parentElement, MoveOperationType.Insert, notesPageDocument.Xnm);
            }
            else if (!processAsExtendedVerse)
            {
                if (!_processedNodes.ContainsKey(notesPageName))
                    _processedNodes.Add(notesPageName, new HashSet<string>());

                if (!_processedNodes[notesPageName].Contains(notePageInfo.Id))
                {
                    suchNoteLink = TryToInsertOrMoveElement(ref oneNoteApp, suchNoteLink.Parent, notePageInfo, _parentElement, MoveOperationType.MoveAndUpdateMetadata, notesPageDocument.Xnm)
                                            .XPathSelectElement("one:T", notesPageDocument.Xnm);
                    _processedNodes[notesPageName].Add(notePageInfo.Id);
                }

                var summaryVersesWeight = GetVerseWeight(suchNoteLink.Value);
                var existingVerseLinksElement = suchNoteLink.Parent.XPathSelectElement("one:OEChildren/one:OE/one:T", notesPageDocument.Xnm);
                if (existingVerseLinksElement != null)
                {
                    OneNoteUtils.NormalizeTextElement(existingVerseLinksElement);
                    summaryVersesWeight = InsertAdditionalVerseLink(ref oneNoteApp, ref existingVerseLinksElement, notePageInfo, notePageContentObjectId, vp, verseWeight, versePosition, summaryVersesWeight);
                }
                else  // значит мы нашли второе упоминание данной ссылки в заметке
                {
                    summaryVersesWeight = InsertSecondVerseLink(ref oneNoteApp, ref suchNoteLink, notePageInfo, notePageContentObjectId, vp, verseWeight, versePosition);                    
                }

                var pageLinkArgs = new List<string>();
                pageLinkArgs.Add(string.Format("{0}={1}", Constants.QueryParameterKey_VerseWeight, summaryVersesWeight));
                if (!string.IsNullOrEmpty(notePageInfo.UniqueId))
                    pageLinkArgs.Add(string.Format("{0}={1}", Constants.QueryParameterKey_NotePageId, notePageInfo.UniqueId));

                var pageLink = OneNoteUtils.GenerateLink(
                                                  ref oneNoteApp, 
                                                  GetVerseLinkTitle(notePageInfo.UniqueTitle, summaryVersesWeight >= Constants.ImportantVerseWeight),
                                                  notePageInfo.Id, notePageInfo.UniqueNoteTitleId,
                                                  pageLinkArgs.ToArray());
                suchNoteLink.Value = pageLink;

                if (suchNoteLink.Parent.XPathSelectElement("one:List", notesPageDocument.Xnm) == null)  // почему то нет номера у строки
                {
                    var listNumberInfo = GetListNumberInfo(_level);
                    suchNoteLink.Parent.AddFirst(new XElement(_nms + "List",
                                                    new XElement(_nms + "Number",
                                                            new XAttribute("numberSequence", listNumberInfo.NumberSequence),
                                                            new XAttribute("numberFormat", listNumberInfo.NumberFormat))));
                }

                if (summaryVersesWeight >= Constants.ImportantVerseWeight)
                {
                    suchNoteLink.Parent.XPathSelectElement("one:List/one:Number", notesPageDocument.Xnm).SetAttributeValue("bold", true);
                }
            }            

            notesPageDocument.WasModified = true;
        }

        private static string GetVerseLinkTitle(string title, bool isImportantVerse)
        {
            if (isImportantVerse)
                return string.Format("<span style='font-weight:bold'>{0}</span>", title);
            else
                return title;
        }

        private decimal InsertSecondVerseLink(ref Application oneNoteApp, ref XElement suchNoteLink, HierarchyElementInfo notePageInfo,
            string notePageContentObjectId, VersePointer vp, decimal verseWeight, XmlCursorPosition versePosition)
        {
            var firstVerseLink = StringUtils.GetAttributeValue(suchNoteLink.Value, "href");
            var firstVersePosition = GetVersePosition(suchNoteLink.Value);
            var firstVerseWeight = GetVerseWeight(suchNoteLink.Value);
            

            firstVerseLink = OneNoteUtils.GetLink(                                
                                GetVerseLinkTitle(
                                    string.Format(Resources.Constants.VerseLinkTemplate, firstVersePosition > versePosition ? 2 : 1),
                                    firstVerseWeight >= Constants.ImportantVerseWeight),
                                firstVerseLink) 
                             + GetExistingMultiVerseString(suchNoteLink.Value);

            var verseLink = OneNoteUtils.GenerateLink(ref oneNoteApp,
                                            GetVerseLinkTitle(
                                                string.Format(Resources.Constants.VerseLinkTemplate, firstVersePosition > versePosition ? 1 : 2), 
                                                verseWeight >= Constants.ImportantVerseWeight),
                                            notePageInfo.Id, notePageContentObjectId,
                                            string.Format("{0}={1}", Constants.QueryParameterKey_VersePosition, versePosition),
                                            string.Format("{0}={1}", Constants.QueryParameterKey_VerseWeight, verseWeight))
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

            return firstVerseWeight.GetValueOrDefault(0) + verseWeight;
        }

        private static XmlCursorPosition? GetVersePosition(string href)
        {
            var s = StringUtils.GetQueryParameterValue(href, Constants.QueryParameterKey_VersePosition);
            if (!string.IsNullOrEmpty(s))
                return new XmlCursorPosition(s);

            return null;
        }

        private static decimal? GetVerseWeight(string href)
        {
            var s = StringUtils.GetQueryParameterValue(href, Constants.QueryParameterKey_VerseWeight);
            if (!string.IsNullOrEmpty(s))
                return decimal.Parse(s);

            return null;
        }

        private decimal InsertAdditionalVerseLink(ref Application oneNoteApp, ref XElement existingVerseLinksElement, HierarchyElementInfo notePageInfo,
            string notePageContentObjectId, VersePointer vp, decimal verseWeight, XmlCursorPosition versePosition, decimal? summaryVersesWeight)
        {
            var links = new List<string>();
            var multiVerseStrings = new List<string>();
            var versesWeight = new List<decimal?>();

            foreach (var existingLink in existingVerseLinksElement.Value.Split(new string[] { ">;" }, StringSplitOptions.None).ToList()
                                        .ConvertAll(link => link + "</a>"))
            {
                var multiVerseString = GetExistingMultiVerseString(existingLink);
                var href = StringUtils.GetAttributeValue(existingLink, "href");
                var weight = GetVerseWeight(href);

                links.Add(href);
                multiVerseStrings.Add(multiVerseString);
                versesWeight.Add(weight);
            }
            

            int insertLinkIndex = 0;
            foreach (var existingLink in links)
            {
                var existingLinkVersePosition = GetVersePosition(existingLink);
                if (existingLinkVersePosition > versePosition)
                    break;
                insertLinkIndex++;
            }

            links.Insert(insertLinkIndex, OneNoteUtils.GetOrGenerateLinkHref(
                                                        ref oneNoteApp, null, notePageInfo.Id, notePageContentObjectId, 
                                                        string.Format("{0}={1}", Constants.QueryParameterKey_VersePosition, versePosition),
                                                        string.Format("{0}={1}", Constants.QueryParameterKey_VerseWeight, verseWeight)));
            multiVerseStrings.Insert(insertLinkIndex, GetMultiVerseString(vp.ParentVersePointer ?? vp));
            versesWeight.Insert(insertLinkIndex, verseWeight);

            existingVerseLinksElement.Value = string.Empty;
            for (int index = 0; index < links.Count; index++)
            {
                existingVerseLinksElement.Value += string.Concat(
                                                        index == 0 ? string.Empty : Resources.Constants.VerseLinksDelimiter,
                                                        OneNoteUtils.GetLink(
                                                                        GetVerseLinkTitle(
                                                                                string.Format(Resources.Constants.VerseLinkTemplate, index + 1),
                                                                                versesWeight[index] >= Constants.ImportantVerseWeight),
                                                                        links[index]),
                                                        multiVerseStrings[index]);
            }          

            return summaryVersesWeight.GetValueOrDefault(0) + verseWeight;
        }

        private XElement SearchExistingNoteLink(ref Application oneNoteApp, XElement rootElement, HierarchyElementInfo notePageInfo, string notePageLink, string notesPageName, XmlNamespaceManager xnm)
        {
            var suchNoteLink = SearchExistingNoteLinkInParent(_parentElement, rootElement, notePageInfo, notePageLink, xnm);

            if (suchNoteLink == null)
            {
                //ищем в других местах
                suchNoteLink = SearchExistingNoteLinkInParent(null, rootElement, notePageInfo, notePageLink, xnm);

                if (suchNoteLink != null)  // нашли в другом месте. Переносим
                {
                    var suchNoteLinkOE = suchNoteLink.Parent;
                    var suchNoteLinkOEChildren = suchNoteLinkOE.Parent;

                    TryToInsertOrMoveElement(ref oneNoteApp, suchNoteLinkOE, notePageInfo, _parentElement, MoveOperationType.MoveAndUpdateMetadata, xnm);

                    if (!_processedNodes.ContainsKey(notesPageName))
                        _processedNodes.Add(notesPageName, new HashSet<string>());

                    if (!_processedNodes[notesPageName].Contains(notePageInfo.Id))
                        _processedNodes[notesPageName].Add(notePageInfo.Id);  // чтоб больше не обрабатывали

                    TryToDeleteTreeStructure(suchNoteLinkOEChildren); // если перенесли последнюю страницу в родителе, рекурсивно смотрим: не надо ли удалять родителей, если они стали пустыми

                    suchNoteLink = SearchExistingNoteLinkInParent(_parentElement, rootElement, notePageInfo, notePageLink, xnm);

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

        private XElement SearchExistingNoteLinkInParent(XElement parentEl, XElement rootElement, HierarchyElementInfo notePageInfo, string notePageLink, XmlNamespaceManager xnm)
        {
            XElement suchNoteLink = null;
            var uniqueNoteId = !string.IsNullOrEmpty(notePageInfo.UniqueId)
                                    ? notePageInfo.UniqueId 
                                    : StringUtils.GetAttributeValue(notePageLink, "page-id");         

            var searchInAllPageString = string.Empty;
            if (parentEl == null)
            {
                searchInAllPageString = ".//";
                parentEl = rootElement;
            }

            if (!string.IsNullOrEmpty(uniqueNoteId))
            {   
                suchNoteLink = parentEl.XPathSelectElement(string.Format("{0}one:OE/one:T[contains(.,'{1}')]", searchInAllPageString, uniqueNoteId), xnm);

                if (suchNoteLink == null)
                {
                    if (string.IsNullOrEmpty(notePageInfo.UniqueId))
                        uniqueNoteId = Uri.EscapeDataString(uniqueNoteId);

                    suchNoteLink = parentEl.XPathSelectElement(
                                        string.Format("{0}one:OE/one:T[contains(translate(.,'abcdefghijklmnopqrstuvwxyz','ABCDEFGHIJKLMNOPQRSTUVWXYZ'),'{1}')]",
                                                    searchInAllPageString, uniqueNoteId.ToUpper()),
                                        xnm);
                }
            }

            return suchNoteLink;
        }

        private void CreateParentTreeStructure(ref Application oneNoteApp, HierarchyElementInfo hierarchyElementInfo, string notebookId, string notesPageName, XmlNamespaceManager xnm)
        {
            if (hierarchyElementInfo.Parent != null)
                CreateParentTreeStructure(ref oneNoteApp, hierarchyElementInfo.Parent, notebookId, notesPageName, xnm);

            var node = _parentElement.XPathSelectElement(
                                    string.Format("one:OE/one:Meta[@name=\"{0}\" and @content=\"{1}\"]", 
                                        Consts.Constants.Key_Id, hierarchyElementInfo.UniqueName), 
                                    xnm);

            if (node == null && hierarchyElementInfo.UniqueName != hierarchyElementInfo.Id)
            {
                node = _parentElement.XPathSelectElement(
                                    string.Format("one:OE/one:Meta[@name=\"{0}\" and @content=\"{1}\"]",
                                        Consts.Constants.Key_Id, hierarchyElementInfo.Id),   // для обратной совместимости, так как раньше и записные книжки индентифицировались по ID
                                    xnm);

                if (node != null)  // если нашли - исправляем                
                    OneNoteUtils.UpdateElementMetaData(node.Parent, Consts.Constants.Key_Id, hierarchyElementInfo.UniqueName, xnm);                
            }

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
                                                    hierarchyElementInfo.Title))
                                    );

                if (displayLevel != null && _level >= displayLevel)
                    node.Add(new XAttribute("collapsed", 1));

                var childNode = new XElement(_nms + "OEChildren");
                node.Add(childNode);

                OneNoteUtils.UpdateElementMetaData(node, Constants.Key_Id, hierarchyElementInfo.UniqueName, xnm);

                TryToInsertOrMoveElement(ref oneNoteApp, node, hierarchyElementInfo, _parentElement, MoveOperationType.Insert, xnm);

                _parentElement = childNode;
            }
            else
            {
                node = node.Parent;

                if (!_processedNodes.ContainsKey(notesPageName))
                    _processedNodes.Add(notesPageName, new HashSet<string>());

                if (!_processedNodes[notesPageName].Contains(hierarchyElementInfo.Id))
                {
                    node.XPathSelectElement("one:T", xnm).Value = hierarchyElementInfo.Title;

                    TryToInsertOrMoveElement(ref oneNoteApp, node, hierarchyElementInfo, _parentElement, MoveOperationType.Move, xnm);

                    _processedNodes[notesPageName].Add(hierarchyElementInfo.Id);
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
                parentHierarchy = notebookHierarchy.Content.Root.XPathSelectElement(string.Format("//one:{0}[@ID=\"{1}\"]", elInfo.GetElementName(), elInfo.Id), xnm).Parent;
            else
                parentHierarchy = OneNoteProxy.Instance.GetHierarchy(ref oneNoteApp, null, HierarchyScope.hsNotebooks).Content.Root;

            var noteLinkInHierarchy = parentHierarchy.XPathSelectElement(string.Format("*[@ID=\"{0}\"]", elInfo.Id), xnm);

            var prevNodesInHierarchy = noteLinkInHierarchy.NodesBeforeSelf();

            if (prevNodesInHierarchy.Count() != 0)
            {
                foreach (var existingLink in parentEl.XPathSelectElements("one:OE", xnm))
                {
                    if (!string.IsNullOrEmpty(elInfo.UniqueId))
                    {
                        var existingTitle = StringUtils.GetText(existingLink.XPathSelectElement("one:T", xnm).Value);
                        if (existingTitle == elInfo.UniqueTitle)                        
                            linkWasFound = true;                        
                        else
                        {
                            if (elInfo.UniqueTitle.CompareTo(existingTitle) < 0)
                                break;

                            prevLink = existingLink;
                        }
                    }
                    else
                    {
                        var existingLinkId = OneNoteUtils.GetElementMetaData(existingLink, Constants.Key_Id, xnm);

                        if (existingLinkId == elInfo.UniqueName)                        
                            linkWasFound = true;                        
                        else
                        {
                            var existingLinkInHierarchy = parentHierarchy.XPathSelectElement(
                                        string.Format("*[@{0}=\"{1}\"]", 
                                            elInfo.Parent != null ? "ID" : "name", 
                                            existingLinkId), 
                                        xnm);
                            if (!prevNodesInHierarchy.Contains(existingLinkInHierarchy))
                                break;

                            prevLink = existingLink;
                        }
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
                OneNoteUtils.UpdateElementMetaData(el, Constants.Key_Id, elInfo.UniqueName, xnm);

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
            var result = string.Empty;

            if (vp.IsMultiVerse)
            {
                if (vp.TopChapter != null && vp.TopVerse != null)
                    result = string.Format("({0}:{1}-{2}:{3})", vp.Chapter, vp.Verse, vp.TopChapter, vp.TopVerse);
                else if (vp.TopChapter != null && vp.IsChapter)
                    result = string.Format("({0}-{1})", vp.Chapter, vp.TopChapter);
                else
                    result = string.Format("(:{0}-{1})", vp.Verse, vp.TopVerse);

                result = FormatMultiVerseString(result);
            }

            return result;
        }        

        private static string GetExistingMultiVerseString(string htmlText)
        {
            var multiVerseString = string.Empty;
            var suchNoteLinkText = string.Empty;

            if (!string.IsNullOrEmpty(htmlText))
                suchNoteLinkText = StringUtils.GetText(htmlText);

            if (!string.IsNullOrEmpty(suchNoteLinkText))
                multiVerseString = Regex.Match(suchNoteLinkText, @"\((\d+)?:\d+\-(\d+:)?\d+\)").Value;

            if (!string.IsNullOrEmpty(multiVerseString))
                multiVerseString = FormatMultiVerseString(multiVerseString);

            return multiVerseString;
        }

        private static string FormatMultiVerseString(string multiVerseString)
        {
            return string.Format("<span style='font-style:italic'>&nbsp;{0}</span>", multiVerseString);
        }
    }
}
