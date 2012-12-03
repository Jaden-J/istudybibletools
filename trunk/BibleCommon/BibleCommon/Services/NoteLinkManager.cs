using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Xml.XPath;
using Microsoft.Office.Interop.OneNote;
using BibleCommon;
using System.Xml;
using BibleCommon.Services;
using BibleCommon.Common;
using BibleCommon.Helpers;
using BibleCommon.Consts;

namespace BibleCommon.Services
{    
    public class NoteLinkManager: IDisposable
    {
        #region Helper classes

        public class NotePageProcessedVerseId
        {
            public string NotePageId { get; set; }
            public string NotesPageName { get; set; }

            public override int GetHashCode()
            {
                return this.NotePageId.GetHashCode() ^ this.NotesPageName.GetHashCode();
            }

            public override bool Equals(object obj)
            {
                NotePageProcessedVerseId otherObj = (NotePageProcessedVerseId)obj;
                return this.NotePageId == otherObj.NotePageId
                    && this.NotesPageName == otherObj.NotesPageName;
            }
        }

        public class ProcessVerseEventArgs : EventArgs
        {
            public bool FoundVerse { get; private set; }
            public bool CancelProcess { get; set; }
            public VersePointer VersePointer { get; set; }
            
            public ProcessVerseEventArgs(bool foundVerse, VersePointer vp)
            {
                FoundVerse = foundVerse;
                VersePointer = vp;
            }            
        }      

        internal class FoundChapterInfo  // в итоге здесь будут только те главы, которые представлены в текущей заметке без стихов
        {
            internal string TextElementObjectId { get; set; }
            internal VersePointerSearchResult VersePointerSearchResult { get; set; }
            internal HierarchySearchManager.HierarchySearchResult HierarchySearchResult { get; set; }
        }

        internal class FoundVerseInfo
        {
            internal int Index { get; set; }
            internal VersePointerSearchResult SearchResult { get; set; }
            internal VerseRecognitionManager.LinkInfo LinkInfo { get; set; }
            internal bool IsInBrackets { get; set; }
            internal bool IsExcluded { get; set; }
            internal int CursorPosition { get; set; }
            internal VersePointerSearchResult GlobalChapterSearchResult { get; set; }
        }

        #endregion        
        
        internal const string DoNotAnalyzeAllPageSymbol = "{}";       
          

        public enum AnalyzeDepth
        {
            OnlyFindVerses = 1,            
            SetVersesLinks = 2,
            Full = 3
        }

       

        internal bool IsExcludedCurrentNotePage { get; set; }
        private Dictionary<NotePageProcessedVerseId, HashSet<SimpleVersePointer>> _notePageProcessedVerses = new Dictionary<NotePageProcessedVerseId, HashSet<SimpleVersePointer>>();  

        private Application _oneNoteApp;
        public NoteLinkManager(Application oneNoteApp)
        {
            _oneNoteApp = oneNoteApp;
        }

        public event EventHandler<ProcessVerseEventArgs> OnNextVerseProcess;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="oneNoteApp"></param>
        /// <param name="sectionGroupId"></param>
        /// <param name="sectionId"></param>
        /// <param name="pageId"></param>
        /// <param name="linkDepth"></param>
        /// <param name="force">Обрабатывать даже ссылки</param>
        public void LinkPageVerses(string sectionGroupId, string sectionId, string pageId, 
            AnalyzeDepth linkDepth, bool force)
        {
            try
            {
                bool wasModified = false;
                OneNoteProxy.PageContent notePageDocument = OneNoteProxy.Instance.GetPageContent(_oneNoteApp, pageId, OneNoteProxy.PageType.NotePage);

                string notePageName = (string)notePageDocument.Content.Root.Attribute("name");

                if (OneNoteUtils.IsRecycleBin(notePageDocument.Content.Root))
                    return;

                bool isSummaryNotesPage = false;

                if (IsSummaryNotesPage(_oneNoteApp, notePageDocument, notePageName))
                {
                    Logger.LogMessage(BibleCommon.Resources.Constants.NoteLinkManagerProcessNotesPage);
                    isSummaryNotesPage = true;
                    if (linkDepth > AnalyzeDepth.SetVersesLinks)
                        linkDepth = AnalyzeDepth.SetVersesLinks;  // на странице заметок только обновляем ссылки
                    notePageDocument.PageType = OneNoteProxy.PageType.NotesPage;  // уточняем тип страницы
                }

                if (notePageName.Contains(DoNotAnalyzeAllPageSymbol))
                    IsExcludedCurrentNotePage = true;

                XElement titleElement = notePageDocument.Content.Root.XPathSelectElement("one:Title/one:OE", notePageDocument.Xnm);
                string pageTitleId = titleElement != null ? (string)titleElement.Attribute("objectID") : null;


                string noteSectionGroupName = null;
                string noteSectionName = null;  
                if (linkDepth > AnalyzeDepth.SetVersesLinks)
                {
                    noteSectionGroupName = OneNoteUtils.GetHierarchyElementName(_oneNoteApp, sectionGroupId);
                    noteSectionName = OneNoteUtils.GetHierarchyElementName(_oneNoteApp, sectionId);
                }

                List<FoundChapterInfo> foundChapters = new List<FoundChapterInfo>();
                var pageIdInfo = new PageIdInfo()
                {
                    SectionGroupName = noteSectionGroupName,
                    SectionName = noteSectionName,
                    PageName = notePageName,
                    PageId = pageId,
                    PageTitleId = pageTitleId
                };

                List<VersePointerSearchResult> pageChaptersSearchResult = ProcessPageTitle(_oneNoteApp, notePageDocument.Content,
                    pageIdInfo, foundChapters, notePageDocument.Xnm, linkDepth, force, isSummaryNotesPage,
                    out wasModified);  // получаем главы текущей страницы, указанные в заголовке (глобальные главы, если больше одной - то не используем их при определении принадлежности только ситхов (:3))                

                List<XElement> processedTextElements = new List<XElement>();

                foreach (XElement oeChildrenElement in notePageDocument.Content.Root.XPathSelectElements("one:Outline/one:OEChildren", notePageDocument.Xnm))
                {
                    if (ProcessTextElements(_oneNoteApp, oeChildrenElement, pageIdInfo, foundChapters, processedTextElements, pageChaptersSearchResult,
                         notePageDocument.Xnm, linkDepth, force, isSummaryNotesPage))
                        wasModified = true;
                }

                if (foundChapters.Count > 0)  // то есть имеются главы, которые указаны в тексте именно как главы, без стихов, и на которые надо делать тоже ссылки
                {
                    ProcessChapters(foundChapters, pageIdInfo, linkDepth, force);                                       
                }

                if (linkDepth >= AnalyzeDepth.Full)                
                    notePageDocument.AddLatestAnalyzeTimeMetaAttribute = true;

                notePageDocument.WasModified = true;
            }
            catch (ProcessAbortedByUserException)
            {                
            }
            catch (Exception ex)
            {
                Logger.LogError(BibleCommon.Resources.Constants.NoteLinkManagerProcessingPageErrors, ex);
            }
        }
        

        /// <summary>
        /// Считается, что страница уже загружена к кэш, потому по pageId мы её быстро достанем
        /// </summary>
        /// <param name="pageId"></param>
        public void SetCursorOnNearestVerse(string pageId)
        {
            try
            {
                OneNoteProxy.PageContent notePageDocument = OneNoteProxy.Instance.GetPageContent(_oneNoteApp, pageId, OneNoteProxy.PageType.NotePage);

                XElement latestTextEl = null;
                int latestVerseStartIndex = -1;

                foreach (var textEl in notePageDocument.Content.Root.XPathSelectElements("//one:OE/one:T", notePageDocument.Xnm))
                {                    
                    var verseStartIndex = textEl.Value.LastIndexOf(Constants.QueryParameter_QuickAnalyze);



                    if (verseStartIndex != -1)
                    {
                        latestTextEl = textEl;
                        latestVerseStartIndex = verseStartIndex;                        
                    }
                }

                if (latestTextEl != null)
                {
                    OneNoteUtils.NormalizeTextElement(latestTextEl);                    
                    if (latestVerseStartIndex != -1)
                    {
                        InsertCursorOnPage(latestTextEl, latestVerseStartIndex);
                    }
                }
            }
            catch (ProcessAbortedByUserException)
            {
            }
            catch (Exception ex)
            {
                Logger.LogError(BibleCommon.Resources.Constants.NoteLinkManagerProcessingPageErrors, ex);
            }
        }

        private void InsertCursorOnPage(XElement textEl, int index)
        {
            XNamespace nms = XNamespace.Get(Constants.OneNoteXmlNs);
            var startLinkSearchString = "<span";
            var endLinkSearchString = "</span>";            

            var startLinkIndex = StringUtils.LastIndexOf(textEl.Value, startLinkSearchString, 0, index);
            var endLinkIndex = textEl.Value.IndexOf(endLinkSearchString, index);

            var link = textEl.Value.Substring(startLinkIndex, endLinkIndex - startLinkIndex + endLinkSearchString.Length);
            var textBefore = textEl.Value.Substring(0, startLinkIndex);
            var textAfter = textEl.Value.Substring(endLinkIndex + endLinkSearchString.Length);            

            textEl.AddAfterSelf(new XElement(nms + "T", 
                                                new XCData(textAfter))
                                            );

            textEl.AddAfterSelf(new XElement(nms + "T", new XAttribute("selected", "all"),
                                                new XCData(string.Empty))
                                            );

            textEl.AddAfterSelf(new XElement(nms + "T", 
                                                new XCData(link))
                                            );

            textEl.ReplaceWith(new XElement(nms + "T",
                                                new XCData(textBefore))
                                            );
        }

        private void ProcessChapters(List<FoundChapterInfo> foundChapters, 
            PageIdInfo notePageId, 
            AnalyzeDepth linkDepth, bool force)
        {
            Logger.LogMessage(BibleCommon.Resources.Constants.NoteLinkManagerChapterProcessing, true, false);

            if (linkDepth >= AnalyzeDepth.Full && !IsExcludedCurrentNotePage)
            {
                foreach (FoundChapterInfo chapterInfo in foundChapters)
                {
                    Logger.LogMessage(".", false, false, false);

                    if (!SettingsManager.Instance.ExcludedVersesLinking)   // иначе мы её обработали сразу же, когда встретили
                    {
                        LinkVerseToNotesPage(_oneNoteApp, chapterInfo.VersePointerSearchResult.VersePointer, true,
                            chapterInfo.HierarchySearchResult.HierarchyObjectInfo,
                            notePageId, chapterInfo.TextElementObjectId, true,
                            SettingsManager.Instance.PageName_Notes, null, SettingsManager.Instance.PageWidth_Notes, 1,
                            (chapterInfo.VersePointerSearchResult.ResultType == VersePointerSearchResult.SearchResultType.ExcludableChapter
                                || chapterInfo.VersePointerSearchResult.ResultType == VersePointerSearchResult.SearchResultType.ExcludableChapterWithoutBookName) ? true : force, false);
                    }

                    if (SettingsManager.Instance.RubbishPage_Use)
                    {
                        if (!SettingsManager.Instance.RubbishPage_ExcludedVersesLinking)   // иначе мы её обработали сразу же, когда встретили
                        {
                            LinkVerseToNotesPage(_oneNoteApp, chapterInfo.VersePointerSearchResult.VersePointer, true,
                                chapterInfo.HierarchySearchResult.HierarchyObjectInfo,
                                notePageId, chapterInfo.TextElementObjectId, false,
                                SettingsManager.Instance.PageName_RubbishNotes, null, SettingsManager.Instance.PageWidth_RubbishNotes, 1,
                                (chapterInfo.VersePointerSearchResult.ResultType == VersePointerSearchResult.SearchResultType.ExcludableChapter
                                    || chapterInfo.VersePointerSearchResult.ResultType == VersePointerSearchResult.SearchResultType.ExcludableChapterWithoutBookName) ? true : force, false);
                        }
                    }
                }
            }

            Logger.LogMessage(string.Empty, false, true, false);
        }

        private List<VersePointerSearchResult> ProcessPageTitle(Application oneNoteApp, XDocument notePageDocument,
            PageIdInfo notePageId,
            List<FoundChapterInfo> foundChapters, XmlNamespaceManager xnm,
            AnalyzeDepth linkDepth, bool force, bool isSummaryNotesPage, out bool wasModified)
        {
            wasModified = false;
            List<VersePointerSearchResult> pageChaptersSearchResult = new List<VersePointerSearchResult>();
            VersePointerSearchResult globalChapterSearchResult = null;   // результат поиска "глобальной" главы 
            VersePointerSearchResult prevResult = null;

            if (ProcessTextElement(oneNoteApp, NotebookGenerator.GetPageTitle(notePageDocument, xnm),
                        notePageId, foundChapters, ref globalChapterSearchResult, ref prevResult, null, linkDepth, force, true, isSummaryNotesPage, searchResult =>
                        {
                            if (VersePointerSearchResult.IsChapter(searchResult.ResultType))
                                pageChaptersSearchResult.Add(searchResult);

                        }))
                wasModified = true;

            return pageChaptersSearchResult;
        }

        private bool IsSummaryNotesPage(Application oneNoteApp, OneNoteProxy.PageContent pageDocument, string pageName)
        {
            string isNotesPage = OneNoteUtils.GetPageMetaData(oneNoteApp, pageDocument.Content.Root, Constants.Key_IsSummaryNotesPage, pageDocument.Xnm);
            if (!string.IsNullOrEmpty(isNotesPage))
            {
                if (bool.Parse(isNotesPage))
                    return true;
            }

            // for back compatibility
            if (pageName.StartsWith(SettingsManager.Instance.PageName_Notes + ".") 
                || pageName.StartsWith(SettingsManager.Instance.PageName_RubbishNotes + "."))
                return true;

            return false;
        }

        private bool ProcessTextElements(Application oneNoteApp, XElement parent,
            PageIdInfo notePageId,
            List<FoundChapterInfo> foundChapters, List<XElement> processedTextElements,
            List<VersePointerSearchResult> pageChaptersSearchResult,
            XmlNamespaceManager xnm, AnalyzeDepth linkDepth, bool force, bool isSummaryNotesPage)
        {
            bool wasModified = false;

            VersePointerSearchResult globalChapterSearchResult;   // результат поиска "глобальной" главы 
            VersePointerSearchResult prevResult;            

            foreach (XElement cellElement in parent.XPathSelectElements(".//one:Table/one:Row/one:Cell", xnm))
            {
                if (ProcessTextElements(oneNoteApp, cellElement, notePageId, foundChapters, processedTextElements, pageChaptersSearchResult,
                        xnm, linkDepth, force, isSummaryNotesPage))
                    wasModified = true;
            }
            
            globalChapterSearchResult = pageChaptersSearchResult.Count == 1 && !pageChaptersSearchResult[0].VersePointer.TopChapter.HasValue 
                                                ? pageChaptersSearchResult[0] : null;   // если в заголовке указана одна глава - то используем её при нахождении только стихов, если же указано несколько - то не используем их
            prevResult = null;            

            foreach (XElement textElement in parent.XPathSelectElements(".//one:T", xnm))
            {
                if (processedTextElements.Contains(textElement))
                {
                    globalChapterSearchResult = pageChaptersSearchResult.Count == 1 ? pageChaptersSearchResult[0] : null;   // если в заголовке указана одна глава - то используем её при нахождении только стихов, если же указано несколько - то не используем их
                    prevResult = null;                    
                    continue;
                }

                if (ProcessTextElement(oneNoteApp, textElement, notePageId, foundChapters,
                                         ref globalChapterSearchResult, ref prevResult, pageChaptersSearchResult, linkDepth, force, false, isSummaryNotesPage, null))
                    wasModified = true;

                processedTextElements.Add(textElement);
            }


            return wasModified;
        }

        private void FireProcessVerseEvent(bool foundVerse, VersePointer vp)
        {
            if (OnNextVerseProcess != null)
            {
                var args = new ProcessVerseEventArgs(foundVerse, vp);
                OnNextVerseProcess(this, args);
                if (args.CancelProcess)
                    throw new ProcessAbortedByUserException();
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="oneNoteApp"></param>
        /// <param name="notebookId"></param>
        /// <param name="textElement"></param>
        /// <param name="noteSectionGroupName"></param>
        /// <param name="noteSectionName"></param>
        /// <param name="notePageName"></param>
        /// <param name="notePageId"></param>
        /// <param name="processedVerses"></param>
        /// <param name="foundChapters"></param>
        /// <param name="globalChapterSearchResult"></param>
        /// <param name="prevResult"></param>
        /// <param name="pageChaptersSearchResult"></param>
        /// <param name="linkDepth"></param>
        /// <param name="force"></param>
        /// <param name="isTitle">анализируем ли заголовок</param>
        /// <param name="onVersePointerFound"></param>
        /// <returns></returns>
        private bool ProcessTextElement(Application oneNoteApp, XElement textElement, PageIdInfo notePageId, List<FoundChapterInfo> foundChapters,
            ref VersePointerSearchResult globalChapterSearchResult, ref VersePointerSearchResult prevResult, 
            List<VersePointerSearchResult> pageChaptersSearchResult,
            AnalyzeDepth linkDepth, bool force, bool isTitle, bool isSummaryNotesPage, Action<VersePointerSearchResult> onVersePointerFound)
        {
            FireProcessVerseEvent(false, null);

            bool result = false;            

            if (textElement != null && !string.IsNullOrEmpty(textElement.Value))
            {
                OneNoteUtils.NormalizeTextElement(textElement);                
                var foundVerses = new List<FoundVerseInfo>();
                var correctVerses = new List<FoundVerseInfo>();

                var globalChapterSearchResultTemp = globalChapterSearchResult;
                var prevResultTemp = prevResult;
                
                IterateTextElementLinks(textElement, globalChapterSearchResult, prevResult, isTitle, isSummaryNotesPage, verseInfo =>
                {
                    foundVerses.Add(verseInfo);
                    return verseInfo.SearchResult.VersePointerHtmlEndIndex;
                });
                
                for (int i = 0; i < foundVerses.Count; i++)
                {
                    if (i < foundVerses.Count - 1)
                    {
                        if (foundVerses[i].SearchResult.VersePointerHtmlEndIndex <= foundVerses[i + 1].SearchResult.VersePointerHtmlStartIndex)   // если не пересекаются стихи
                            correctVerses.Add(foundVerses[i]);
                    }
                    else
                        correctVerses.Add(foundVerses[i]);
                }                               

                IterateTextElementLinks(textElement, globalChapterSearchResult, prevResult, isTitle, isSummaryNotesPage, verseInfo =>
                {
                    int cursorPosition = verseInfo.CursorPosition;

                    if (correctVerses.Any(cv => cv.Index == verseInfo.Index))
                    {
                        var processVerseResult = ProcessFoundVerse(cursorPosition, ref verseInfo, textElement, notePageId, foundChapters, pageChaptersSearchResult, linkDepth, force, isTitle, onVersePointerFound);
                        cursorPosition = processVerseResult.CursorPosition;
                        if (processVerseResult.WasModified)
                            result = processVerseResult.WasModified;

                        globalChapterSearchResultTemp = verseInfo.GlobalChapterSearchResult;
                        prevResultTemp = verseInfo.SearchResult;                        
                    }
                    else
                        cursorPosition = verseInfo.SearchResult.VersePointerHtmlEndIndex; 

                    return cursorPosition;
                });

                globalChapterSearchResult = globalChapterSearchResultTemp;
                prevResult = prevResultTemp;
            }

            return result;
        }



        private ProcessFoundVerseResult ProcessFoundVerse(int cursorPosition, ref FoundVerseInfo verseInfo, XElement textElement,
            PageIdInfo notePageId, List<FoundChapterInfo> foundChapters, List<VersePointerSearchResult> pageChaptersSearchResult, 
            AnalyzeDepth linkDepth, bool force, bool isTitle, Action<VersePointerSearchResult> onVersePointerFound)
        {
            bool wasModified = false;
            string textElementValue = textElement.Value;

            FireProcessVerseEvent(true, verseInfo.SearchResult.VersePointer);

            if (onVersePointerFound != null)
                onVersePointerFound(verseInfo.SearchResult);
            
            HierarchySearchManager.HierarchySearchResult hierarchySearchResult;            

            if (!verseInfo.LinkInfo.IsLink
                || (verseInfo.LinkInfo.LinkType == VerseRecognitionManager.LinkInfo.LinkTypeEnum.LinkAfterQuickAnalyze && (linkDepth >= AnalyzeDepth.Full || force))
                || (verseInfo.LinkInfo.LinkType == VerseRecognitionManager.LinkInfo.LinkTypeEnum.LinkAfterFullAnalyze && force)
                || IsExtendedVerse(ref verseInfo))
            {
                string textElementObjectId = (string)textElement.Parent.Attribute("objectID");

                bool needToQueueIfChapter;
                textElementValue = ProcessVerse(_oneNoteApp, verseInfo.SearchResult,                                            
                                            textElementValue,
                                            notePageId, textElementObjectId,
                                            linkDepth, verseInfo.GlobalChapterSearchResult, pageChaptersSearchResult,
                                            verseInfo.LinkInfo, verseInfo.IsInBrackets, verseInfo.IsExcluded, force,
                                            out cursorPosition, out hierarchySearchResult, out needToQueueIfChapter);

                var tempVerseInfo = verseInfo;

                if (verseInfo.SearchResult.ResultType == VersePointerSearchResult.SearchResultType.SingleVerseOnly)  // то есть нашли стих, а до этого значит была скорее всего просто глава!
                {
                    FoundChapterInfo chapterInfo = foundChapters.FirstOrDefault(fch =>
                            fch.VersePointerSearchResult.ResultType != VersePointerSearchResult.SearchResultType.ExcludableChapter
                            && fch.VersePointerSearchResult.ResultType != VersePointerSearchResult.SearchResultType.ExcludableChapterWithoutBookName
                            && fch.VersePointerSearchResult.VersePointer.Book.Name == tempVerseInfo.SearchResult.VersePointer.Book.Name
                            && IsNumberInRange(tempVerseInfo.SearchResult.VersePointer.Chapter.Value, fch.VersePointerSearchResult.VersePointer.Chapter.Value, fch.VersePointerSearchResult.VersePointer.TopChapter));                                

                    if (chapterInfo != null)
                        foundChapters.Remove(chapterInfo);
                }
                else if (VersePointerSearchResult.IsChapter(verseInfo.SearchResult.ResultType) && needToQueueIfChapter)
                {
                    if (hierarchySearchResult.ResultType == HierarchySearchManager.HierarchySearchResultType.Successfully
                        && hierarchySearchResult.HierarchyStage == HierarchySearchManager.HierarchyStage.Page)
                    {
                        if (!foundChapters.Exists(fch =>
                                (fch.VersePointerSearchResult.ResultType == VersePointerSearchResult.SearchResultType.ExcludableChapter
                                    || fch.VersePointerSearchResult.ResultType == VersePointerSearchResult.SearchResultType.ExcludableChapterWithoutBookName)
                                && fch.VersePointerSearchResult.VersePointer.Book.Name == tempVerseInfo.SearchResult.VersePointer.Book.Name
                                && IsNumberInRange(tempVerseInfo.SearchResult.VersePointer.Chapter.Value, fch.VersePointerSearchResult.VersePointer.Chapter.Value, fch.VersePointerSearchResult.VersePointer.TopChapter)))
                        {
                            foundChapters.Add(new FoundChapterInfo()
                            {
                                TextElementObjectId = textElementObjectId,
                                VersePointerSearchResult = verseInfo.SearchResult,
                                HierarchySearchResult = hierarchySearchResult
                            });
                        }
                    }
                }

                if (textElement.Value != textElementValue)
                {
                    textElement.Value = textElementValue;
                    wasModified = true;
                }
            }
            else
                cursorPosition = verseInfo.SearchResult.VersePointerHtmlEndIndex;

            return new ProcessFoundVerseResult()
                       {
                           CursorPosition = cursorPosition,
                           WasModified = wasModified
                       };
        }

        /// <summary>
        /// Является ли стих случаем, когда у нас был "Ин 1:2". Мы проанализировали. А потом добавили "-22"
        /// </summary>
        /// <param name="verseInfo"></param>
        /// <returns></returns>
        private bool IsExtendedVerse(ref FoundVerseInfo verseInfo)
        {
            if ((verseInfo.LinkInfo.LinkType == VerseRecognitionManager.LinkInfo.LinkTypeEnum.LinkAfterFullAnalyze
                 || verseInfo.LinkInfo.LinkType == VerseRecognitionManager.LinkInfo.LinkTypeEnum.LinkAfterQuickAnalyze) && verseInfo.SearchResult.VersePointer.IsMultiVerse)
            {   
                var link = verseInfo.SearchResult.TextElement.Value.Substring(verseInfo.SearchResult.VersePointerStartIndex, verseInfo.SearchResult.VersePointerEndIndex - verseInfo.SearchResult.VersePointerStartIndex);
                var indexOfLink = link.IndexOf("</a>");
                if (indexOfLink != -1)
                {
                    link = link.Substring(0, indexOfLink);
                    if (!link.Contains('-'))   // то есть вроде как бы IsMultiVerse, но при этом нет тире внутри самой ссылки
                    {
                        if (verseInfo.LinkInfo.LinkType == VerseRecognitionManager.LinkInfo.LinkTypeEnum.LinkAfterFullAnalyze)
                            verseInfo.LinkInfo.ExtendedVerse = true;  // помечаем, чтобы потом не анализировать первый стих, который уже анализировали
                        return true; 
                    }
                }
            }

            return false;
        }

        private void IterateTextElementLinks(XElement textElement, VersePointerSearchResult globalChapterSearchResult, VersePointerSearchResult prevResult, 
            bool isTitle, bool isSummaryNotesPage, Func<FoundVerseInfo, int> verseAction)
        {

            string localChapterName = string.Empty;    // имя главы в пределах данного абзаца. например, действительно только для девятки в "Откр 5:7,9"
            int cursorPosition = StringUtils.GetNextIndexOfDigit(textElement.Value, null);

            int verseIndex = 0;
            while (cursorPosition > -1)
            {
                try
                {
                    int number;
                    int textBreakIndex;
                    int htmlBreakIndex;
                    VerseRecognitionManager.LinkInfo linkInfo;
                    bool isInBrackets;
                    bool isExcluded;
                    if (VerseRecognitionManager.CanProcessAtNumberPosition(textElement, cursorPosition,
                        out number, out textBreakIndex, out htmlBreakIndex, out linkInfo, out isInBrackets, out isExcluded))
                    {
                        var searchResult = VerseRecognitionManager.GetValidVersePointer(textElement,
                            cursorPosition, textBreakIndex - 1, number,
                            globalChapterSearchResult,
                            localChapterName, prevResult, linkInfo.IsLink, isInBrackets, isTitle);

                        if (searchResult.ResultType != VersePointerSearchResult.SearchResultType.Nothing && isSummaryNotesPage)
                            if (searchResult.VersePointer != null && searchResult.VersePointer.IsMultiVerse)  // если находимся на странице сводной заметок и нашли мультивёрс ссылку (например :4-7) - то такие ссылки не обрабатываем
                            {
                                searchResult.ResultType = VersePointerSearchResult.SearchResultType.Nothing;
                                cursorPosition = searchResult.VersePointerHtmlEndIndex;
                            }

                        if (searchResult.ResultType != VersePointerSearchResult.SearchResultType.Nothing)
                        {
                            if (searchResult.ResultType == VersePointerSearchResult.SearchResultType.ChapterOnlyAtStartString
                                || searchResult.ResultType == VersePointerSearchResult.SearchResultType.ChapterAndVerseAtStartString)
                            {
                                globalChapterSearchResult = searchResult;
                            }

                            cursorPosition = verseAction(new FoundVerseInfo()
                                                             {
                                                                 Index = verseIndex++,
                                                                 SearchResult = searchResult,
                                                                 LinkInfo = linkInfo,
                                                                 IsInBrackets = isInBrackets,
                                                                 IsExcluded = isExcluded,
                                                                 CursorPosition = cursorPosition,
                                                                 GlobalChapterSearchResult = globalChapterSearchResult
                                                             });


                            localChapterName = searchResult.ChapterName;   // всегда запоминаем
                            prevResult = searchResult;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Logger.LogError(ex);
                }

                cursorPosition = StringUtils.GetNextIndexOfDigit(textElement.Value, cursorPosition);
            }
        }
        

     


        /// <summary>
        /// 
        /// </summary>
        /// <param name="oneNoteApp"></param>
        /// <param name="notebookId"></param>
        /// <param name="searchResult"></param>
        /// <param name="processedVerses"></param>
        /// <param name="textToChange"></param>
        /// <param name="textElementValue"></param>
        /// <param name="noteSectionGroupName"></param>
        /// <param name="noteSectionName"></param>
        /// <param name="notePageName"></param>
        /// <param name="notePageId"></param>
        /// <param name="notePageContentObjectId"></param>
        /// <param name="linkDepth"></param>
        /// <param name="globalChapterSearchResult"></param>
        /// <param name="pageChaptersSearchResult">главы страницы, в числе которых могут быть главы в квадрытных скобках (исключаемые главы)</param>
        /// <param name="linkInfo"></param>        
        /// <param name="force"></param>        
        /// <param name="newEndVerseIndex"></param>
        /// <param name="hierarchySearchResult"></param>
        /// <returns></returns>
        private string ProcessVerse(Application oneNoteApp, VersePointerSearchResult searchResult,
            string textElementValue, PageIdInfo notePageId, string notePageContentObjectId,
            AnalyzeDepth linkDepth, VersePointerSearchResult globalChapterSearchResult, List<VersePointerSearchResult> pageChaptersSearchResult,
            VerseRecognitionManager.LinkInfo linkInfo, bool isInBrackets, bool isExcluded, bool force, 
            out int newEndVerseIndex, out HierarchySearchManager.HierarchySearchResult hierarchySearchResult, out bool needToQueueIfChapter)
        {
            hierarchySearchResult = new HierarchySearchManager.HierarchySearchResult() { ResultType = HierarchySearchManager.HierarchySearchResultType.NotFound };
            var localHierarchySearchResult = new HierarchySearchManager.HierarchySearchResult() { ResultType = HierarchySearchManager.HierarchySearchResultType.NotFound };
            needToQueueIfChapter = true;  // по умолчанию - анализируем главы в самом конце            

            int startVerseNameIndex = searchResult.VersePointerStartIndex;
            int endVerseNameIndex = searchResult.VersePointerEndIndex;

            newEndVerseIndex = endVerseNameIndex;

            bool wasModifiedOnLinkCorrection;
            if (!CorrectTextToChangeBoundary(ref textElementValue, linkInfo,
                              ref startVerseNameIndex, ref endVerseNameIndex, out wasModifiedOnLinkCorrection))
            {
                newEndVerseIndex = searchResult.VersePointerHtmlEndIndex; // потому что это же значение мы присваиваем, если стоит !force и встретили гиперссылку                
                return textElementValue;
            }

            if (!wasModifiedOnLinkCorrection)
            {             
                if (searchResult.VersePointerHtmlStartIndex != searchResult.VersePointerStartIndex)
                {
                    if (StringUtils.GetChar(textElementValue, searchResult.VersePointerHtmlStartIndex) == VerseRecognitionManager.ChapterVerseDelimiter
                        && StringUtils.GetChar(textElementValue, searchResult.VersePointerHtmlStartIndex + 1) == '<')  // случай типа "<span lang=ru>:</span><span lang=en-US>12</span>"
                    {
                        startVerseNameIndex = searchResult.VersePointerHtmlStartIndex;
                        endVerseNameIndex = searchResult.VersePointerHtmlEndIndex;
                    }
                }
            }

            if (searchResult.VersePointer.IsChapter)
                needToQueueIfChapter = !NeedToForceAnalyzeChapter(searchResult); // нужно ли анализировать главу сразу же, а не в самом конце             

            List<VersePointer> allIncludedVersesExceptFirst = null;

            #region linking main notes page            

            List<VersePointer> verses = new List<VersePointer>() { searchResult.VersePointer };

            if (SettingsManager.Instance.ExpandMultiVersesLinking && searchResult.VersePointer.IsMultiVerse)
            {
                allIncludedVersesExceptFirst = searchResult.VersePointer.GetAllIncludedVersesExceptFirst(oneNoteApp,
                                                    new GetAllIncludedVersesExceptFirstArgs() { BibleNotebookId = SettingsManager.Instance.NotebookId_Bible });
                verses.AddRange(allIncludedVersesExceptFirst);
            }


            bool processAsExtendedVerse = linkDepth >= AnalyzeDepth.Full && !force && linkInfo.ExtendedVerse;            

            bool first = true;
            var processedVerses = new List<SimpleVersePointer>();

            foreach (VersePointer vp in verses)
            {
                if (processedVerses.Contains(vp.ToSimpleVersePointer()))
                    continue;                
                
                if (TryLinkVerseToNotesPage(oneNoteApp, vp, searchResult.ResultType,
                        notePageId, notePageContentObjectId, linkDepth,
                        !SettingsManager.Instance.UseDifferentPagesForEachVerse || (vp.IsChapter && !needToQueueIfChapter), SettingsManager.Instance.ExcludedVersesLinking,
                        SettingsManager.Instance.PageName_Notes, null, SettingsManager.Instance.PageWidth_Notes, 1,
                        globalChapterSearchResult, pageChaptersSearchResult,
                        isInBrackets, isExcluded, force, !needToQueueIfChapter, processAsExtendedVerse,
                        out localHierarchySearchResult, ref processedVerses, hsr =>
                        {
                            if (first)                            
                                Logger.LogMessage("{0}: {1}", 
                                    searchResult.VersePointer.IsChapter ? BibleCommon.Resources.Constants.ProcessChapter : BibleCommon.Resources.Constants.ProcessVerse, 
                                    searchResult.VersePointer.OriginalVerseName);                                                            
                        }))
                {
                    if (first)
                    {                        
                        if (linkDepth >= AnalyzeDepth.SetVersesLinks)
                        {
                            string textToChange;
                            if (!string.IsNullOrEmpty(searchResult.VerseString))
                                textToChange = searchResult.VerseString;
                            else if (VersePointerSearchResult.IsChapter(searchResult.ResultType))
                                textToChange = searchResult.ChapterName;
                            else
                                textToChange = searchResult.VersePointer.OriginalVerseName;

                             hierarchySearchResult = localHierarchySearchResult;

                            var prevLinkText = linkInfo.IsLink ? textElementValue.Substring(startVerseNameIndex, endVerseNameIndex - startVerseNameIndex) : null;

                            var additionalParams = new List<string>();                            

                            if (linkDepth == AnalyzeDepth.SetVersesLinks)
                            {
                                if (!linkInfo.IsLink
                                    || prevLinkText.Contains(Consts.Constants.QueryParameter_QuickAnalyze)
                                    || (wasModifiedOnLinkCorrection && linkInfo.ExtendedVerse))
                                    additionalParams.Add(Consts.Constants.QueryParameter_QuickAnalyze);

                                if (linkInfo.ExtendedVerse 
                                    || (!string.IsNullOrEmpty(prevLinkText) && prevLinkText.Contains(Consts.Constants.QueryParameter_ExtendedVerse)))
                                    additionalParams.Add(Consts.Constants.QueryParameter_ExtendedVerse);
                            }

                            string prevStyle = string.Empty;
                            if (linkInfo.IsLink)
                                prevStyle = StringUtils.GetAttributeValue(prevLinkText, "style");

                            string link = OneNoteUtils.GetOrGenerateHref(oneNoteApp, textToChange,                                
                                            localHierarchySearchResult.HierarchyObjectInfo.VerseInfo.ObjectHref,
                                            localHierarchySearchResult.HierarchyObjectInfo.PageId, localHierarchySearchResult.HierarchyObjectInfo.VerseContentObjectId, additionalParams.ToArray());

                            link = string.Format("<span style='font-weight:normal;{1}'>{0}</span>", link, prevStyle);

                            var htmlTextBefore = textElementValue.Substring(0, startVerseNameIndex);
                            var htmlTextAfter = textElementValue.Substring(endVerseNameIndex);                            
                            var needToAddSpace = false;
                            if (searchResult.VerseStringStartsWithSpace.GetValueOrDefault(false))
                            {
                                var textBefore = StringUtils.GetText(htmlTextBefore);
                                if (!textBefore.EndsWith(" "))
                                    needToAddSpace = true;
                            }

                            textElementValue = string.Concat(
                                htmlTextBefore,
                                needToAddSpace ? " " : string.Empty,
                                link,
                                htmlTextAfter);

                            newEndVerseIndex = startVerseNameIndex + link.Length + (needToAddSpace ? 1 : 0);
                            searchResult.VersePointerHtmlEndIndex = newEndVerseIndex;
                        }
                    }
                }

                if (SettingsManager.Instance.UseDifferentPagesForEachVerse && !vp.IsChapter)  // для каждого стиха своя страница
                {                    
                    string notesPageName = GetDefaultNotesPageName(
                                hierarchySearchResult.HierarchyObjectInfo.AdditionalObjectsIds.ContainsKey(vp)
                                    ? (VerseNumber?)hierarchySearchResult.HierarchyObjectInfo.AdditionalObjectsIds[vp].VerseNumber
                                    : hierarchySearchResult.HierarchyObjectInfo.VerseNumber);
                    TryLinkVerseToNotesPage(oneNoteApp, vp, searchResult.ResultType,
                        notePageId, notePageContentObjectId, linkDepth,
                        true, SettingsManager.Instance.ExcludedVersesLinking,
                        notesPageName, SettingsManager.Instance.PageName_Notes, SettingsManager.Instance.PageWidth_Notes, 2,
                        globalChapterSearchResult, pageChaptersSearchResult,
                        isInBrackets, isExcluded, force, !needToQueueIfChapter, processAsExtendedVerse, out localHierarchySearchResult, ref processedVerses, null);
                }

                first = false;

                System.Windows.Forms.Application.DoEvents();
            }           

            #endregion           

            #region linking rubbish notes pages

            if (SettingsManager.Instance.RubbishPage_Use)
            {
                List<VersePointer> rubbishVerses = new List<VersePointer>() { searchResult.VersePointer };

                if (SettingsManager.Instance.RubbishPage_ExpandMultiVersesLinking && searchResult.VersePointer.IsMultiVerse)
                {
                    if (allIncludedVersesExceptFirst == null)
                        allIncludedVersesExceptFirst = searchResult.VersePointer.GetAllIncludedVersesExceptFirst(oneNoteApp,
                                                            new GetAllIncludedVersesExceptFirstArgs() { BibleNotebookId = SettingsManager.Instance.NotebookId_Bible });
                    rubbishVerses.AddRange(allIncludedVersesExceptFirst);
                }
                
                foreach (VersePointer vp in rubbishVerses)
                {
                    TryLinkVerseToNotesPage(oneNoteApp, vp, searchResult.ResultType, 
                        notePageId, notePageContentObjectId, linkDepth,
                        false, SettingsManager.Instance.RubbishPage_ExcludedVersesLinking, 
                        SettingsManager.Instance.PageName_RubbishNotes, null, SettingsManager.Instance.PageWidth_RubbishNotes, 1,
                        globalChapterSearchResult, pageChaptersSearchResult,
                        isInBrackets, isExcluded, force, !needToQueueIfChapter, processAsExtendedVerse, out localHierarchySearchResult, ref processedVerses, null);                

                    System.Windows.Forms.Application.DoEvents();
                }
            }

            #endregion

            return textElementValue;
        }

        private bool NeedToForceAnalyzeChapter(VersePointerSearchResult searchResult)
        {
            return searchResult.VersePointer.IsMultiVerse
                    || (searchResult.ResultType != VersePointerSearchResult.SearchResultType.ChapterOnlyAtStartString
                        && searchResult.ResultType != VersePointerSearchResult.SearchResultType.ChapterAndVerseAtStartString);
        }

        internal static string GetDefaultNotesPageName(VerseNumber? verseNumber)
        {
            if (verseNumber.HasValue && SettingsManager.Instance.UseDifferentPagesForEachVerse)
                return string.Format("{1} {2}", SettingsManager.Instance.PageName_Notes, verseNumber, 
                    verseNumber.Value.IsMultiVerse ? BibleCommon.Resources.Constants.Verses : BibleCommon.Resources.Constants.Verse);

            return SettingsManager.Instance.PageName_Notes;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="oneNoteApp"></param>
        /// <param name="vp"></param>
        /// <param name="resultType"></param>
        /// <param name="processedVerses"></param>
        /// <param name="noteSectionGroupName"></param>
        /// <param name="noteSectionName"></param>
        /// <param name="notePageName"></param>
        /// <param name="notePageId"></param>
        /// <param name="notePageTitleId"></param>
        /// <param name="notePageContentObjectId"></param>
        /// <param name="linkDepth"></param>
        /// <param name="createLinkToNotesPage">Необходримо ли создавать сслыку на страницу сводной заметок. Если например мы обрабатываем RubbishPage, то такая ссылка не нужна</param>
        /// <param name="excludedVersesLinking">Привязываем ли стихи, даже если они входятв главу, являющуюся ExcludableChapter</param>
        /// <param name="globalChapterSearchResult"></param>
        /// <param name="pageChaptersSearchResult"></param>
        /// <param name="isInBrackets"></param>
        /// <param name="force"></param>
        /// <param name="hierarchySearchResult"></param>
        /// <param name="onHierarchyElementFound"></param>
        /// <param name="notesPageLevel">1, 2 or 3</param>
        /// <returns></returns>
        private bool TryLinkVerseToNotesPage(Application oneNoteApp, VersePointer vp,
            VersePointerSearchResult.SearchResultType resultType, 
            PageIdInfo notePageId, string notePageContentObjectId, AnalyzeDepth linkDepth,
            bool createLinkToNotesPage, bool excludedVersesLinking, 
            string notesPageName, string notesParentPageName, int notesPageWidth, int notesPageLevel,
            VersePointerSearchResult globalChapterSearchResult, List<VersePointerSearchResult> pageChaptersSearchResult,
            bool isInBrackets, bool isExcluded, bool force, bool forceAnalyzeChapter, bool processAsExtendedVerse, out HierarchySearchManager.HierarchySearchResult hierarchySearchResult,
            ref List<SimpleVersePointer> processedVerses,
            Action<HierarchySearchManager.HierarchySearchResult> onHierarchyElementFound)
        {
            hierarchySearchResult = HierarchySearchManager.GetHierarchyObject(
                                                oneNoteApp, SettingsManager.Instance.NotebookId_Bible, vp, HierarchySearchManager.FindVerseLevel.AllVerses);
            if (hierarchySearchResult.ResultType == HierarchySearchManager.HierarchySearchResultType.Successfully)
            {
                if (hierarchySearchResult.HierarchyStage == HierarchySearchManager.HierarchyStage.ContentPlaceholder
                    || hierarchySearchResult.HierarchyStage == HierarchySearchManager.HierarchyStage.Page)
                {
                    if (hierarchySearchResult.HierarchyObjectInfo.PageId != notePageId.PageId)
                    {
                        if (onHierarchyElementFound != null)
                            onHierarchyElementFound(hierarchySearchResult);

                        bool isChapter = VersePointerSearchResult.IsChapter(resultType);

                        if ((isChapter ^ vp.IsChapter))  // данные расходятся. что-то тут не чисто. запишем варнинг и ничего делать не будем
                        {
                            Logger.LogWarning("Invalid verse result: '{0}' - '{1}'", vp.OriginalVerseName, resultType);
                        }
                        else
                        {
                            if ((!isChapter || excludedVersesLinking || forceAnalyzeChapter) && linkDepth >= AnalyzeDepth.Full)   // главы сразу не обрабатываем - вдруг есть стихи этих глав в текущей заметке. Вот если нет - тогда потом и обработаем. Но если у нас стоит excludedVersesLinking, то сразу обрабатываем
                            {
                                bool canContinue = true;

                                if (!excludedVersesLinking)  // иначе всё равно привязываем
                                {
                                    if (isExcluded || IsExcludedCurrentNotePage)
                                        canContinue = false;

                                    if (canContinue)
                                    {
                                        if (VersePointerSearchResult.IsVerse(resultType))
                                        {
                                            if (globalChapterSearchResult != null)
                                            {
                                                if (globalChapterSearchResult.ResultType == VersePointerSearchResult.SearchResultType.ChapterAndVerseAtStartString)
                                                {
                                                    if (!isInBrackets)
                                                    {
                                                        if (globalChapterSearchResult.VersePointer.IsMultiVerse)
                                                            if (globalChapterSearchResult.VersePointer.IsInVerseRange(vp.Verse.GetValueOrDefault(0)))
                                                                canContinue = false;    // если указан уже диапазон, а далее идут пояснения, то не отмечаем их заметками
                                                    }
                                                }
                                                //else
                                                //{
                                                //    if (globalChapterSearchResult.ResultType == VersePointerSearchResult.SearchResultType.ExcludableChapter)
                                                //    {
                                                //        if (!isInBrackets)
                                                //        {
                                                //            canContinue = false;
                                                //        }
                                                //    }
                                                //}
                                            }
                                        }
                                    }

                                    if (canContinue)
                                    {
                                        if (pageChaptersSearchResult != null)
                                        {
                                            if (!isInBrackets)
                                            {
                                                if (pageChaptersSearchResult.Any(pcsr =>
                                                {
                                                    return (pcsr.ResultType == VersePointerSearchResult.SearchResultType.ExcludableChapter
                                                                || pcsr.ResultType == VersePointerSearchResult.SearchResultType.ExcludableChapterWithoutBookName)
                                                            && pcsr.VersePointer.Book.Name == vp.Book.Name
                                                            && IsNumberInRange(vp.Chapter.Value, pcsr.VersePointer.Chapter.Value, pcsr.VersePointer.TopChapter);
                                                }))
                                                {
                                                    canContinue = false;  // то есть среди исключаемых глав есть текущая
                                                }
                                            }
                                        }
                                    }
                                }

                                if (canContinue)
                                {
                                    var verses = LinkVerseToNotesPage(oneNoteApp, vp, isChapter,
                                        hierarchySearchResult.HierarchyObjectInfo,
                                        notePageId,
                                        notePageContentObjectId, createLinkToNotesPage, notesPageName, notesParentPageName, notesPageWidth, notesPageLevel, force, processAsExtendedVerse);

                                    if (processedVerses != null)
                                        processedVerses.AddRange(verses);
                                }
                            }

                            return true;
                        }
                    }
                }
            }

            return false;
        }

        private bool IsNumberInRange(int number, int bottom, int? top)
        {
            if (top.HasValue)            
                return number >= bottom && number <= top;            
            else
                return number == bottom;
        }


        
        /// <summary>
        /// 
        /// </summary>
        /// <param name="textElementValue"></param>
        /// <param name="linkInfo"></param>
        /// <param name="startVerseNameIndex"></param>
        /// <param name="endVerseNameIndex"></param>
        /// <returns>false - если помимо библейской ссылки, в гиперссылке содержится и другой текст. Не обрабатываем такие ссылки</returns>
        private bool CorrectTextToChangeBoundary(ref string textElementValue, VerseRecognitionManager.LinkInfo linkInfo, ref int startVerseNameIndex, ref int endVerseNameIndex, out bool wasModified)
        {
            wasModified = false;

            if (linkInfo.IsLink)
            {
                string beginSearchString = "<a ";
                string endSearchString = "</a>";
                int linkStartIndex = StringUtils.LastIndexOf(textElementValue, beginSearchString, 0, endVerseNameIndex);
                if (linkStartIndex != -1)
                {
                    var linkEndIndexInside = textElementValue.IndexOf(endSearchString, linkStartIndex);
                    var linkEndIndexOutside = textElementValue.IndexOf(endSearchString, endVerseNameIndex);

                    if (linkEndIndexInside != -1 && linkEndIndexInside != linkEndIndexOutside)  // то есть ссылка окончилась внутри. Тогда просто переносим </a> на конец библейской ссылки
                    {
                        var textBeforeEndLink = textElementValue.Substring(0, linkEndIndexInside);
                        var textAfterEndLinkButBeforeEndVerse = textElementValue.Substring(linkEndIndexInside + endSearchString.Length, endVerseNameIndex - linkEndIndexInside - endSearchString.Length);
                        var textAfter = textElementValue.Substring(endVerseNameIndex);
                        textElementValue = string.Concat(textBeforeEndLink, textAfterEndLinkButBeforeEndVerse, endSearchString, textAfter);
                        endVerseNameIndex -= endSearchString.Length;
                        linkEndIndexOutside = textElementValue.IndexOf(endSearchString, endVerseNameIndex);
                    }

                    if (linkEndIndexOutside != -1)
                    {
                        int startVerseNameIndexTemp = linkStartIndex;
                        int endVerseNameIndexTemp = linkEndIndexOutside + endSearchString.Length;

                        string textBefore = startVerseNameIndexTemp < startVerseNameIndex 
                                                ? StringUtils.GetText(textElementValue.Substring(startVerseNameIndexTemp, startVerseNameIndex - startVerseNameIndexTemp))
                                                : StringUtils.GetText(textElementValue.Substring(startVerseNameIndex, startVerseNameIndexTemp - startVerseNameIndex));
                        if (string.IsNullOrEmpty(textBefore.Trim()))                        
                        {
                            string textAfter = StringUtils.GetText(textElementValue.Substring(endVerseNameIndex, endVerseNameIndexTemp - endVerseNameIndex));
                            if (string.IsNullOrEmpty(textAfter.Trim()))
                            {
                                startVerseNameIndex = startVerseNameIndexTemp;
                                endVerseNameIndex = endVerseNameIndexTemp;
                                wasModified = true;
                                return true;
                            }
                        }
                        
                        return false; // иначе помимо библеской ссылки есть и другой текст в этой гиперссылке. 
                    }
                }
            }


            return true;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="oneNoteApp"></param>
        /// <param name="vp"></param>
        /// <param name="isChapter"></param>
        /// <param name="processedVerses"></param>
        /// <param name="verseHierarchyObjectInfo"></param>
        /// <param name="noteSectionGroupName"></param>
        /// <param name="noteSectionName"></param>
        /// <param name="notePageName"></param>
        /// <param name="notePageId"></param>
        /// <param name="notePageTitleId"></param>
        /// <param name="notePageContentObjectId"></param>
        /// <param name="createLinkToNotesPage">Необходримо ли создавать сслыку на страницу сводной заметок. Если например мы обрабатываем RubbishPage, то такая ссылка не нужна</param>
        /// <param name="notesPageName">название страницы "Сводная заметок"</param>
        /// <param name="force"></param>
        private List<SimpleVersePointer> LinkVerseToNotesPage(Application oneNoteApp, VersePointer vp, bool isChapter,
            HierarchySearchManager.HierarchyObjectInfo verseHierarchyObjectInfo,
            PageIdInfo notePageId, string notePageContentObjectId, bool createLinkToNotesPage,
            string notesPageName, string notesParentPageName, int notesPageWidth, int notesPageLevel, bool force, bool processAsExtendedVerse)
        {
            string biblePageName = verseHierarchyObjectInfo.PageName;

            string notesPageId = null;
            bool pageWasCreated = false;
            bool rowWasAdded = false;
            try
            {
                notesPageId = OneNoteProxy.Instance.GetNotesPageId(oneNoteApp,
                    verseHierarchyObjectInfo.SectionId,
                    verseHierarchyObjectInfo.PageId, biblePageName, notesPageName, out pageWasCreated, notesParentPageName, notesPageLevel);                    
            }
            catch (Exception ex)
            {
                Logger.LogError(ex);
            }

            if (!string.IsNullOrEmpty(notesPageId))
            {
                string targetContentObjectId = NotesPageManager.UpdateNotesPage(oneNoteApp, this, vp, isChapter, verseHierarchyObjectInfo,
                        notePageId, notesPageId, notePageContentObjectId, 
                        notesPageName, notesPageWidth, force, processAsExtendedVerse, out rowWasAdded);

                if (createLinkToNotesPage && (pageWasCreated || rowWasAdded))
                {
                    OneNoteProxy.PageContent versePageDocument = OneNoteProxy.Instance.GetPageContent(oneNoteApp, verseHierarchyObjectInfo.PageId, OneNoteProxy.PageType.Bible);       

                    string link = string.Format("<font size='2pt'>{0}</font>",
                                    OneNoteUtils.GenerateHref(oneNoteApp, SettingsManager.Instance.PageName_Notes, notesPageId, null)); // здесь всегда передаём null, так как в частых случаях он и так null, потому что страница в кэше, и в OneNote она ещё не обновлялась (то есть идентификаторы ещё не проставлены). Так как эти идентификаторы проставятся в самом конце, то и ссылки обновим в конце.

                    bool wasModified = false;

                    if (isChapter)
                    {
                        if (SetLinkToNotesPageForChapter(versePageDocument.Content, link, versePageDocument.Xnm))
                            wasModified = true;
                    }
                    else
                    {
                        if (SetLinkToNotesPageForVerse(versePageDocument.Content, link, vp, verseHierarchyObjectInfo, versePageDocument.Xnm))
                            wasModified = true;
                    }

                    if (wasModified)
                    {
                        versePageDocument.WasModified = true;
                        
                        OneNoteProxy.Instance.AddProcessedBiblePageWithUpdatedLinksToNotesPages(verseHierarchyObjectInfo.SectionId, verseHierarchyObjectInfo.PageId, biblePageName, vp.GetChapterPointer());

                        OneNoteProxy.Instance.AddProcessedVerseOnBiblePageWithUpdatedLinksToNotesPages(vp, verseHierarchyObjectInfo.VerseNumber);  // добавляем только стихи, отмеченные на "Сводной заметок"
                    }                    
                }

                var key = new NotePageProcessedVerseId() { NotePageId = notePageId.PageId, NotesPageName = notesPageName };
                return AddNotePageProcessedVerse(key, vp, verseHierarchyObjectInfo.VerseNumber);                
            }

            return new List<SimpleVersePointer>();
        }

        private bool SetLinkToNotesPageForVerse(XDocument pageDocument, string link, VersePointer vp,
            HierarchySearchManager.HierarchyObjectInfo verseHierarchyObjectInfo, XmlNamespaceManager xnm)
        {
            bool result = false;

            //находим ячейку для заметок стиха
            XElement contentObject = pageDocument.XPathSelectElement(string.Format("//one:OE[@objectID = '{0}']",
                verseHierarchyObjectInfo.VerseContentObjectId), xnm);
            if (contentObject == null)
            {
                Logger.LogError("{0} '{1}'", BibleCommon.Resources.Constants.NoteLinkManagerVerseCellNotFound,  vp.OriginalVerseName);
            }
            else
            {
                XNode cellForNotesPage = contentObject.Parent.Parent.NextNode;
                XElement textElement = cellForNotesPage.XPathSelectElement("one:OEChildren/one:OE/one:T", xnm);

                //if (textElement.Value == string.Empty)  // лучше обновлять ссылку на страницу заметок, так как зачастую она вначале бывает неточной (только на страницу)
                {
                    textElement.Value = link;
                    result = true;
                }
            }

            return result;
        }


        /// <summary>
        /// Возвращает элемент - ссылку на страницу сводную заметок на ряд для главы
        /// </summary>
        /// <param name="pageDocument"></param>
        /// <param name="xnm"></param>
        /// <returns></returns>
        internal static XElement GetChapterNotesPageLink(XDocument pageDocument, XmlNamespaceManager xnm)
        {
            XElement notesLinkElement = pageDocument.Root.XPathSelectElement("one:Outline/one:OEChildren", xnm);

            if (notesLinkElement != null && notesLinkElement.Nodes().Count() == 1)   // похоже на правду                        
            {
                notesLinkElement = notesLinkElement.XPathSelectElement(string.Format("one:OE/one:T[contains(.,'>{0}<')]",
                    SettingsManager.Instance.PageName_Notes), xnm);
            }
            else notesLinkElement = null;

            return notesLinkElement;
        }

        private bool SetLinkToNotesPageForChapter(XDocument pageDocument, string link, XmlNamespaceManager xnm)
        {
            bool result = false;

            XElement notesLinkElement = GetChapterNotesPageLink(pageDocument, xnm);                

            if (notesLinkElement == null)
            {
                XElement titleElement = pageDocument.Root.XPathSelectElement("one:Title", xnm);

                XNamespace nms = XNamespace.Get(Constants.OneNoteXmlNs);
                titleElement.AddAfterSelf(new XElement(nms + "Outline",
                                 new XElement(nms + "Position",
                                        new XAttribute("x", "360.0"),
                                        new XAttribute("y", "14.40000057220459"),
                                        new XAttribute("z", "1")),
                                 new XElement(nms + "OEChildren",
                                   new XElement(nms + "OE",
                                       new XElement(nms + "T",
                                           new XCData(link))))));

                result = true;
            }
            else
            {
                notesLinkElement.Value = link;
                result = true;
            }

            return result;
        }        

       

       

        #region helper methods

        public List<SimpleVersePointer> AddNotePageProcessedVerse(NotePageProcessedVerseId verseId, VersePointer vp, VerseNumber? verseNumber)
        {
            if (!_notePageProcessedVerses.ContainsKey(verseId))
            {
                _notePageProcessedVerses.Add(verseId, new HashSet<SimpleVersePointer>());
            }

            var svp = vp.ToSimpleVersePointer();
            if (verseNumber.HasValue)
                svp.VerseNumber = verseNumber.Value;

            var result = svp.GetAllVerses();

            if (!_notePageProcessedVerses[verseId].Contains(svp))   // отслеживаем обработанные стихи для каждой из страниц сводной заметок
            {               
                result.ForEach(v => _notePageProcessedVerses[verseId].Add(v));                
            }

            return result;
        }

        public bool ContainsNotePageProcessedVerse(NotePageProcessedVerseId verseId, VersePointer vp)
        {
            if (!_notePageProcessedVerses.ContainsKey(verseId))
            {
                _notePageProcessedVerses.Add(verseId, new HashSet<SimpleVersePointer>());
            }

            return _notePageProcessedVerses[verseId].Contains(vp.ToSimpleVersePointer());
        }
      
        #endregion


        public void Dispose()
        {
            _oneNoteApp = null;
        }
    }
}
