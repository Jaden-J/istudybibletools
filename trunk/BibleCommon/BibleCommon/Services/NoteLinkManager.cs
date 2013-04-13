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
using BibleCommon.Handlers;
using System.IO;
using BibleCommon.Providers;

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
            internal BibleSearchResult HierarchySearchResult { get; set; }
            internal bool IsImportantChapter { get; set; }
            internal XmlCursorPosition ChapterPosition { get; set; }
            internal decimal ChapterWeight { get; set; }
        }

        public class FoundVerseInfo
        {
            public int Index { get; set; }
            public VersePointerSearchResult SearchResult { get; set; }
            public VerseRecognitionManager.LinkInfo LinkInfo { get; set; }
            public bool IsInBrackets { get; set; }
            public bool IsExcluded { get; set; }
            public int CursorPosition { get; set; }
            public VersePointerSearchResult GlobalChapterSearchResult { get; set; }
            public bool IsImportantVerse { get; set; }
            public XmlCursorPosition VersePosition { get; set; }
        }

        

        #endregion        
        
        public enum AnalyzeDepth
        {
            OnlyFindVerses = 1,            
            SetVersesLinks = 2,
            Full = 3
        }

        public bool AnalyzeAllPages { get; set; }
        public FoundVerseInfo LastAnalyzedVerse { get; set; }

        internal bool IsExcludedCurrentNotePage { get; set; }
        private Dictionary<NotePageProcessedVerseId, HashSet<SimpleVersePointer>> _notePageProcessedVerses = new Dictionary<NotePageProcessedVerseId, HashSet<SimpleVersePointer>>();
        private Dictionary<NotePageProcessedVerseId, HashSet<SimpleVersePointer>> _notePageProcessedVersesForOldProvider = new Dictionary<NotePageProcessedVerseId, HashSet<SimpleVersePointer>>();  

        private Application _oneNoteApp;
        private NotesPagesProviderManager _notesPagesProviderManager;
        private bool _pageContentWasChanged = false;


        public NoteLinkManager(Application oneNoteApp)
        {
            _oneNoteApp = oneNoteApp;
            _notesPagesProviderManager = new NotesPagesProviderManager();
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
        /// <param name="doNotAnalyze">В названии элементов иерархии встретился символ '{}'. То есть не анализируем такую страницу.</param>
        public void LinkPageVerses(string notebookId, string pageId, AnalyzeDepth linkDepth, bool force, bool? doNotAnalyze)
        {
            try
            {   
                _notesPagesProviderManager.ForceUpdateProvider = AnalyzeAllPages && force && linkDepth >= AnalyzeDepth.Full;

                bool wasModified = false;
                OneNoteProxy.PageContent notePageDocument = OneNoteProxy.Instance.GetPageContent(ref _oneNoteApp, pageId, OneNoteProxy.PageType.NotePage, PageInfo.piBasic, true);

                string notePageName = (string)notePageDocument.Content.Root.Attribute("name");

                if (OneNoteUtils.IsRecycleBin(notePageDocument.Content.Root))
                    return;

                bool isSummaryNotesPage = false;

                if (IsSummaryNotesPage(ref _oneNoteApp, notePageDocument, notePageName))
                {
                    Logger.LogMessageParams(BibleCommon.Resources.Constants.NoteLinkManagerProcessNotesPage);
                    isSummaryNotesPage = true;
                    if (linkDepth > AnalyzeDepth.SetVersesLinks)
                        linkDepth = AnalyzeDepth.SetVersesLinks;  // на странице заметок только обновляем ссылки
                    notePageDocument.PageType = OneNoteProxy.PageType.NotesPage;  // уточняем тип страницы
                }

                if (doNotAnalyze.GetValueOrDefault(false) || StringUtils.IndexOfAny(notePageName, Constants.DoNotAnalyzeSymbol1, Constants.DoNotAnalyzeSymbol2) > -1)
                    IsExcludedCurrentNotePage = true;               

                HierarchyElementInfo notePageHierarchyInfo;

                if (linkDepth > AnalyzeDepth.SetVersesLinks)
                    notePageHierarchyInfo = GetPageHierarchyInfo(notebookId, notePageDocument, pageId, notePageName, true);
                else
                    notePageHierarchyInfo = GetPageHierarchyInfo(notebookId, notePageDocument, pageId, notePageName, false);

                var foundChapters = new List<FoundChapterInfo>();                

                List<VersePointerSearchResult> pageChaptersSearchResult = ProcessPageTitle(ref _oneNoteApp, notePageDocument.Content,
                    notePageHierarchyInfo, ref foundChapters, notePageDocument.Xnm, linkDepth, force, isSummaryNotesPage,
                    out wasModified);  // получаем главы текущей страницы, указанные в заголовке (глобальные главы, если больше одной - то не используем их при определении принадлежности только ситхов (:3))                

                List<XElement> processedTextElements = new List<XElement>();

                var dictionaryMetaData = OneNoteUtils.GetElementMetaData(notePageDocument.Content.Root, Constants.Key_EmbeddedDictionaries, notePageDocument.Xnm);
                if (dictionaryMetaData != null)  // значит мы анализируем страницу словаря
                {
                    var dictionaryName = dictionaryMetaData.Split(new char[] { ',' })[0];

                    foreach (XElement oeParent in notePageDocument.Content.Root.XPathSelectElements("one:Outline/one:OEChildren/one:OE/one:Table/one:Row", notePageDocument.Xnm))
                    {
                        var termEl = oeParent.XPathSelectElement("one:Cell[1]/one:OEChildren/one:OE/one:T", notePageDocument.Xnm);
                        var termName = StringUtils.GetText(termEl.Value);
                        var termElId = (string)termEl.Parent.Attribute("objectID");

                        notePageHierarchyInfo.ManualId = string.Format("{{{0}}}", Uri.EscapeUriString(string.Format("{0}_|_{1}", dictionaryName, termName)));
                        notePageHierarchyInfo.UniqueTitle = termName;
                        notePageHierarchyInfo.UniqueNoteTitleId = termElId;

                        if (ProcessTextElements(ref _oneNoteApp, oeParent, notePageHierarchyInfo, ref foundChapters, processedTextElements, pageChaptersSearchResult,
                             notePageDocument.Xnm, linkDepth, force, isSummaryNotesPage))
                            wasModified = true;
                    }
                }
                else
                {
                    foreach (XElement oeChildrenElement in notePageDocument.Content.Root.XPathSelectElements("one:Outline/one:OEChildren", notePageDocument.Xnm))
                    {
                        if (ProcessTextElements(ref _oneNoteApp, oeChildrenElement, notePageHierarchyInfo, ref foundChapters, processedTextElements, pageChaptersSearchResult,
                             notePageDocument.Xnm, linkDepth, force, isSummaryNotesPage))
                            wasModified = true;
                    }
                }

                if (foundChapters.Count > 0)  // то есть имеются главы, которые указаны в тексте именно как главы, без стихов, и на которые надо делать тоже ссылки
                {
                    ProcessChapters(foundChapters, notePageHierarchyInfo, linkDepth, force);                                       
                }

                notePageDocument.WasModified = true;

                if (linkDepth >= AnalyzeDepth.Full && _pageContentWasChanged)
                {
                    notePageDocument.AddLatestAnalyzeTimeMetaAttribute = true;

                    Logger.LogMessageParams(Resources.Constants.UpdatingPageInOneNote);
                    System.Windows.Forms.Application.DoEvents();
                    OneNoteProxy.Instance.CommitModifiedPage(ref _oneNoteApp, notePageDocument, false);
                }
                else
                    OneNoteProxy.Instance.RemovePageContentFromCache(pageId, PageInfo.piBasic);
            }
            catch (ProcessAbortedByUserException)
            {                
            }
            catch (Exception ex)
            {
                Logger.LogError(BibleCommon.Resources.Constants.NoteLinkManagerProcessingPageErrors, ex);
            }
        }

        private HierarchyElementInfo GetPageHierarchyInfo(string notebookId, OneNoteProxy.PageContent notePageDocument, string notePageId, string notePageName, bool loadFullHierarchy)
        {
            XElement titleElement = notePageDocument.Content.Root.XPathSelectElement("one:Title/one:OE", notePageDocument.Xnm);
            string pageTitleId = titleElement != null ? (string)titleElement.Attribute("objectID") : null;

            var result = new HierarchyElementInfo()
                {
                    Title = notePageName,
                    Id = notePageId,
                    Type = HierarchyElementType.Page,
                    PageTitleId = pageTitleId,
                    NotebookId = notebookId
                };

            if (loadFullHierarchy) 
            {
                result.SyncPageId = OneNoteUtils.GetElementMetaData(notePageDocument.Content.Root, Constants.Key_SyncId, notePageDocument.Xnm);
                if (result.SyncPageId == null)
                {
                    result.SyncPageId = notePageId;
                    OneNoteUtils.UpdateElementMetaData(notePageDocument.Content.Root, Constants.Key_SyncId, result.SyncPageId, notePageDocument.Xnm);                    
                }                

                var fullNotebookHierarchy = OneNoteProxy.Instance.GetHierarchy(ref _oneNoteApp, notebookId, HierarchyScope.hsPages, false);
                LoadHierarchyElementParent(notebookId, fullNotebookHierarchy, ref result);
            }

            return result;
        }

        private void LoadHierarchyElementParent(string notebookId, OneNoteProxy.HierarchyElement fullNotebookHierarchy, ref HierarchyElementInfo elementInfo)
        {
            var el = fullNotebookHierarchy.Content.Root.XPathSelectElement(
                                string.Format("//one:{0}[@ID=\"{1}\"]", elementInfo.GetElementName(), elementInfo.Id), fullNotebookHierarchy.Xnm);

            if (el == null)
                throw new Exception(string.Format("Can not find hierarchyElement '{0}' of type '{1}' in notebook '{2}'", 
                                elementInfo.Id, elementInfo.Type, notebookId));


            if (el.Parent != null)
            {
                var parentId = (string)el.Parent.Attribute("ID");
                var parentType = (HierarchyElementType)Enum.Parse(typeof(HierarchyElementType), el.Parent.Name.LocalName);

                string parentName;
                string parentTitle;
                if (parentType == HierarchyElementType.Notebook)
                {
                    parentTitle = (string)el.Parent.Attribute("nickname");
                    parentName = (string)el.Parent.Attribute("name");

                    if (string.IsNullOrEmpty(parentTitle))
                        parentTitle = parentName;
                }
                else
                {
                    parentName = (string)el.Parent.Attribute("name");
                    parentTitle = parentName;
                }

                var parent = new HierarchyElementInfo() { Id = parentId, Title = parentTitle, Name = parentName, Type = parentType, NotebookId = notebookId };                
                LoadHierarchyElementParent(notebookId, fullNotebookHierarchy, ref parent);
                elementInfo.Parent = parent;
            }
        }       
        
        public void SetCursorOnNearestVerse(FoundVerseInfo verseInfo)
        {
            try
            {
                if (verseInfo != null)
                    InsertCursorOnPage(verseInfo.SearchResult.TextElement, verseInfo.SearchResult.VersePointerHtmlEndIndex);
            }
            catch (ProcessAbortedByUserException)
            {
            }
            catch (Exception ex)
            {
                Logger.LogError(BibleCommon.Resources.Constants.NoteLinkManagerProcessingPageErrors, ex);
            }
        }

        private void InsertCursorOnPage(XElement textEl, int endIndex)
        {
            XNamespace nms = XNamespace.Get(Constants.OneNoteXmlNs);
            var startLinkSearchString = "<a";
            var endLinkSearchString = "</a>";            

            var startLinkIndex = StringUtils.LastIndexOf(textEl.Value, startLinkSearchString, 0, endIndex);
            var endLinkIndex = textEl.Value.IndexOf(endLinkSearchString, startLinkIndex);         

            if (startLinkIndex != -1 && endLinkIndex != -1)
            {
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
        }

        private void ProcessChapters(List<FoundChapterInfo> foundChapters, 
            HierarchyElementInfo notePageId, 
            AnalyzeDepth linkDepth, bool force)
        {
            Logger.LogMessage(BibleCommon.Resources.Constants.NoteLinkManagerChapterProcessing, true, false);

            if (linkDepth >= AnalyzeDepth.Full && !IsExcludedCurrentNotePage)
            {
                foreach (var chapterInfo in foundChapters)
                {
                    Logger.LogMessage(".", false, false, false);

                    if (!SettingsManager.Instance.ExcludedVersesLinking)   // иначе мы её обработали сразу же, когда встретили
                    {
                        LinkVerseToNotesPage(ref _oneNoteApp, chapterInfo.VersePointerSearchResult.VersePointer, chapterInfo.ChapterWeight, 
                            chapterInfo.ChapterPosition, true,
                            chapterInfo.HierarchySearchResult.HierarchyObjectInfo,
                            notePageId, chapterInfo.TextElementObjectId, true,
                            NotesPageType.Chapter, SettingsManager.Instance.PageWidth_Notes, 1, chapterInfo.IsImportantChapter,
                            (chapterInfo.VersePointerSearchResult.ResultType == VersePointerSearchResult.SearchResultType.ExcludableChapter
                                || chapterInfo.VersePointerSearchResult.ResultType == VersePointerSearchResult.SearchResultType.ExcludableChapterWithoutBookName) ? true : force, false);
                    }

                    if (SettingsManager.Instance.RubbishPage_Use)
                    {
                        if (!SettingsManager.Instance.RubbishPage_ExcludedVersesLinking)   // иначе мы её обработали сразу же, когда встретили
                        {
                            LinkVerseToNotesPage(ref _oneNoteApp, chapterInfo.VersePointerSearchResult.VersePointer, chapterInfo.ChapterWeight, 
                                chapterInfo.ChapterPosition, true,
                                chapterInfo.HierarchySearchResult.HierarchyObjectInfo,
                                notePageId, chapterInfo.TextElementObjectId, false,
                                NotesPageType.Detailed, SettingsManager.Instance.PageWidth_RubbishNotes, 1, chapterInfo.IsImportantChapter,
                                (chapterInfo.VersePointerSearchResult.ResultType == VersePointerSearchResult.SearchResultType.ExcludableChapter
                                    || chapterInfo.VersePointerSearchResult.ResultType == VersePointerSearchResult.SearchResultType.ExcludableChapterWithoutBookName) ? true : force, false);
                        }
                    }
                }
            }

            Logger.LogMessage(string.Empty, false, true, false);
        }

        private List<VersePointerSearchResult> ProcessPageTitle(ref Application oneNoteApp, XDocument notePageDocument,
            HierarchyElementInfo notePageId,
            ref List<FoundChapterInfo> foundChapters, XmlNamespaceManager xnm,
            AnalyzeDepth linkDepth, bool force, bool isSummaryNotesPage, out bool wasModified)
        {
            wasModified = false;
            List<VersePointerSearchResult> pageChaptersSearchResult = new List<VersePointerSearchResult>();
            VersePointerSearchResult globalChapterSearchResult = null;   // результат поиска "глобальной" главы 
            VersePointerSearchResult prevResult = null;

            if (ProcessTextElement(NotebookGenerator.GetPageTitle(notePageDocument, xnm),
                        notePageId, ref foundChapters, ref globalChapterSearchResult, ref prevResult, null, linkDepth, force, true, isSummaryNotesPage, searchResult =>
                        {
                            if (VersePointerSearchResult.IsChapter(searchResult.ResultType))
                                pageChaptersSearchResult.Add(searchResult);

                        }))
                wasModified = true;

            pageChaptersSearchResult = pageChaptersSearchResult.Where(sr => sr.VersePointer.IsChapter).ToList();  // так как могли быть указаны стихи типа "2Ин 8"

            return pageChaptersSearchResult;
        }

        private bool IsSummaryNotesPage(ref Application oneNoteApp, OneNoteProxy.PageContent pageDocument, string pageName)
        {
            string isNotesPage = OneNoteUtils.GetElementMetaData(pageDocument.Content.Root, Constants.Key_IsSummaryNotesPage, pageDocument.Xnm);
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

        private bool ProcessTextElements(ref Application oneNoteApp, XElement parent,
            HierarchyElementInfo notePageId,
            ref List<FoundChapterInfo> foundChapters, List<XElement> processedTextElements,
            List<VersePointerSearchResult> pageChaptersSearchResult,
            XmlNamespaceManager xnm, AnalyzeDepth linkDepth, bool force, bool isSummaryNotesPage)
        {
            bool wasModified = false;

            VersePointerSearchResult globalChapterSearchResult;   // результат поиска "глобальной" главы 
            VersePointerSearchResult prevResult;            

            foreach (XElement cellElement in parent.XPathSelectElements(".//one:Table/one:Row/one:Cell", xnm))
            {
                if (ProcessTextElements(ref oneNoteApp, cellElement, notePageId, ref foundChapters, processedTextElements, pageChaptersSearchResult,
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

                if (ProcessTextElement(textElement, notePageId, ref foundChapters,
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
        private bool ProcessTextElement(XElement textElement, HierarchyElementInfo notePageId, ref List<FoundChapterInfo> foundChapters,
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
                var correctVerses = new Dictionary<int, FoundVerseInfo>();

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
                            correctVerses.Add(foundVerses[i].Index, foundVerses[i]);
                    }
                    else
                        correctVerses.Add(foundVerses[i].Index, foundVerses[i]);
                }

                var foundChaptersLocal = foundChapters;
                IterateTextElementLinks(textElement, globalChapterSearchResult, prevResult, isTitle, isSummaryNotesPage, verseInfo =>
                {
                    int cursorPosition = verseInfo.CursorPosition;

                    if (correctVerses.ContainsKey(verseInfo.Index))
                    {
                        var processVerseResult = ProcessFoundVerse(cursorPosition, ref verseInfo, textElement, notePageId, ref foundChaptersLocal, pageChaptersSearchResult, linkDepth, force, isTitle, onVersePointerFound);
                        cursorPosition = processVerseResult.CursorPosition;
                        if (processVerseResult.WasModified)
                            result = processVerseResult.WasModified;

                        globalChapterSearchResultTemp = verseInfo.GlobalChapterSearchResult;
                        prevResultTemp = verseInfo.SearchResult;                        
                    }
                    else
                        cursorPosition = verseInfo.SearchResult.VersePointerHtmlEndIndex;


                    System.Windows.Forms.Application.DoEvents();

                    return cursorPosition;
                });
                foundChapters = foundChaptersLocal;

                globalChapterSearchResult = globalChapterSearchResultTemp;
                prevResult = prevResultTemp;
            }

            return result;
        }



        private ProcessFoundVerseResult ProcessFoundVerse(int cursorPosition, ref FoundVerseInfo verseInfo, XElement textElement,
            HierarchyElementInfo notePageId, ref List<FoundChapterInfo> foundChapters, List<VersePointerSearchResult> pageChaptersSearchResult, 
            AnalyzeDepth linkDepth, bool force, bool isTitle, Action<VersePointerSearchResult> onVersePointerFound)
        {
            bool wasModified = false;
            string textElementValue = textElement.Value;

            FireProcessVerseEvent(true, verseInfo.SearchResult.VersePointer);

            if (onVersePointerFound != null)
                onVersePointerFound(verseInfo.SearchResult);

            BibleSearchResult hierarchySearchResult;            

            if (!verseInfo.LinkInfo.IsLink
                || (verseInfo.LinkInfo.LinkType == VerseRecognitionManager.LinkInfo.LinkTypeEnum.LinkAfterQuickAnalyze && (linkDepth >= AnalyzeDepth.Full || force))
                || (verseInfo.LinkInfo.LinkType == VerseRecognitionManager.LinkInfo.LinkTypeEnum.LinkAfterFullAnalyze && force)
                || IsExtendedVerse(ref verseInfo))
            {
                string textElementObjectId = (string)textElement.Parent.Attribute("objectID");

                bool needToQueueIfChapter;
                decimal verseWeight;
                XmlCursorPosition versePosition;
                textElementValue = ProcessVerse(ref _oneNoteApp, verseInfo.SearchResult,                                            
                                            textElementValue,
                                            notePageId, textElementObjectId,
                                            linkDepth, verseInfo.GlobalChapterSearchResult, pageChaptersSearchResult,
                                            verseInfo.LinkInfo, verseInfo.IsInBrackets, verseInfo.IsExcluded, verseInfo.IsImportantVerse, force,
                                            out cursorPosition, out hierarchySearchResult, out needToQueueIfChapter, out verseWeight, out versePosition);

                verseInfo.VersePosition = versePosition;

                if (LastAnalyzedVerse == null || LastAnalyzedVerse.VersePosition <= verseInfo.VersePosition)
                    LastAnalyzedVerse = verseInfo;

                var tempVerseInfo = verseInfo;
                if (verseInfo.SearchResult.ResultType == VersePointerSearchResult.SearchResultType.SingleVerseOnly)  // то есть нашли стих без главы, а до этого значит была скорее всего просто глава!
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
                    if (hierarchySearchResult.ResultType == BibleHierarchySearchResultType.Successfully
                        && hierarchySearchResult.HierarchyStage == BibleHierarchyStage.Page)  // здесь мы ещё раз проверяем, что это глава. Так как мог быть стих "2Ин 8"
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
                                HierarchySearchResult = hierarchySearchResult,
                                IsImportantChapter = verseInfo.IsImportantVerse,
                                ChapterWeight = verseWeight,
                                ChapterPosition = versePosition
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
                    bool isImportantVerse;
                    if (VerseRecognitionManager.CanProcessAtNumberPosition(textElement, cursorPosition, 
                        out number, out textBreakIndex, out htmlBreakIndex, out linkInfo, out isInBrackets, out isExcluded, out isImportantVerse))
                    {
                        isImportantVerse = isImportantVerse || isTitle;
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
                                                                 GlobalChapterSearchResult = globalChapterSearchResult,
                                                                 IsImportantVerse = isImportantVerse
                                                             });


                            localChapterName = string.Format("{0} {1}", 
                                searchResult.VersePointer.OriginalBookName, searchResult.VersePointer.TopChapter.GetValueOrDefault(searchResult.VersePointer.Chapter.Value));  // всегда запоминаем  

                            //ранее было "localChapterName = searchResult.ChapterName;". Пришлось изменить, чтобы понимать правильную главу для стиха Ин 2:3 в "Ин 1:1-2:2,3".

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
        private string ProcessVerse(ref Application oneNoteApp, VersePointerSearchResult searchResult,
            string textElementValue, HierarchyElementInfo notePageId, string notePageContentObjectId,
            AnalyzeDepth linkDepth, VersePointerSearchResult globalChapterSearchResult, List<VersePointerSearchResult> pageChaptersSearchResult,
            VerseRecognitionManager.LinkInfo linkInfo, bool isInBrackets, bool isExcluded, bool isImportantVerse, bool force,
            out int newEndVerseIndex, out BibleSearchResult hierarchySearchResult, out bool needToQueueIfChapter, 
            out decimal verseWeight, out XmlCursorPosition versePosition)
        {
            hierarchySearchResult = new BibleSearchResult() { ResultType = BibleHierarchySearchResultType.NotFound };
            var localHierarchySearchResult = new BibleSearchResult() { ResultType = BibleHierarchySearchResultType.NotFound };
            needToQueueIfChapter = true;  // по умолчанию - анализируем главы в самом конце            
            verseWeight = 1;
            versePosition = new XmlCursorPosition((IXmlLineInfo)searchResult.TextElement);

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

            List<VersePointer> allSubVerses = null;

            #region linking main notes page            

            var verses = new List<VersePointer>();

            if (searchResult.VersePointer.IsMultiVerse)
            {
                allSubVerses = searchResult.VersePointer.GetAllVerses(ref oneNoteApp,
                                                    new GetAllIncludedVersesExceptFirstArgs() { BibleNotebookId = SettingsManager.Instance.NotebookId_Bible });

                verseWeight = (decimal)1 / (decimal)allSubVerses.Count;

                if (SettingsManager.Instance.ExpandMultiVersesLinking)
                    verses.AddRange(allSubVerses);
            }            

            if (verses.Count == 0)
                verses.Add(searchResult.VersePointer);

            if (isImportantVerse)
                verseWeight = Constants.ImportantVerseWeight;
            else
                verseWeight = decimal.Round(verseWeight, 2, MidpointRounding.AwayFromZero);

            bool processAsExtendedVerse = linkDepth >= AnalyzeDepth.Full && !force && linkInfo.ExtendedVerse;            

            bool first = true;
            var processedVerses = new List<SimpleVersePointer>();            

            foreach (VersePointer vp in verses)
            {
                if (processedVerses.Contains(vp.ToSimpleVersePointer()))
                    continue;                
                
                if (TryLinkVerseToNotesPage(ref oneNoteApp, vp, verseWeight, searchResult.ResultType, versePosition,
                        notePageId, notePageContentObjectId, linkDepth,
                        !SettingsManager.Instance.UseDifferentPagesForEachVerse || (vp.IsChapter && !needToQueueIfChapter), SettingsManager.Instance.ExcludedVersesLinking,
                        NotesPageType.Chapter, SettingsManager.Instance.PageWidth_Notes, 1, 
                        globalChapterSearchResult, pageChaptersSearchResult,
                        isInBrackets, isExcluded, isImportantVerse, force, !needToQueueIfChapter, processAsExtendedVerse,
                        out localHierarchySearchResult, ref processedVerses, hsr =>
                        {
                            if (first)                            
                                Logger.LogMessageParams("{0}: {1}", 
                                    searchResult.VersePointer.IsChapter ? BibleCommon.Resources.Constants.ProcessChapter : BibleCommon.Resources.Constants.ProcessVerse, 
                                    searchResult.VersePointer.OriginalVerseName);                                                            
                        }))
                {
                    if (first)
                    {                        
                        if (linkDepth >= AnalyzeDepth.SetVersesLinks)
                        {
                            if (vp.ParentVersePointer != null)
                            {
                                var parentVerse = vp.ParentVersePointer;
                                BibleHierarchySearchProvider.CheckVerseForExisting(ref oneNoteApp, ref parentVerse);   // чтобы, если например Иуд 1-4, изменить, как надо                                
                            }

                            string textToChange;
                            if (!string.IsNullOrEmpty(searchResult.VerseString))
                                textToChange = searchResult.VerseString;
                            else if (vp.IsChapter)
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


                            var linkHref = SettingsManager.Instance.UseProxyLinksForBibleVerses
                                                ? OpenBibleVerseHandler.GetCommandUrlStatic(vp.ParentVersePointer ?? vp, SettingsManager.Instance.ModuleShortName)
                                                : localHierarchySearchResult.HierarchyObjectInfo.VerseInfo.ObjectHref;

                            string link = OneNoteUtils.GetOrGenerateLink(ref oneNoteApp, textToChange, linkHref,
                                            localHierarchySearchResult.HierarchyObjectInfo.PageId, 
                                            localHierarchySearchResult.HierarchyObjectInfo.VerseContentObjectId, additionalParams.ToArray());

                            link = string.Format("<span style='font-weight:normal;{1}'>{0}</span>", link, prevStyle);

                            var htmlTextBefore = textElementValue.Substring(0, startVerseNameIndex);
                            var htmlTextAfter = textElementValue.Substring(endVerseNameIndex);                            
                            var needToAddSpace = false;
                            if (searchResult.VerseStringStartsWithSpace.GetValueOrDefault(false) && !vp.WasChangedVerseAsOneChapteredBook)
                            {
                                var textBefore = StringUtils.GetText(htmlTextBefore);
                                if (!textBefore.EndsWith(" "))
                                    needToAddSpace = true;
                            }

                            CheckIfLinkIsChanged(linkInfo, textElementValue, startVerseNameIndex, endVerseNameIndex, linkHref, link);                            

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
                    TryLinkVerseToNotesPage(ref oneNoteApp, vp, verseWeight, searchResult.ResultType, versePosition,
                        notePageId, notePageContentObjectId, linkDepth,
                        true, SettingsManager.Instance.ExcludedVersesLinking,
                        NotesPageType.Verse, SettingsManager.Instance.PageWidth_Notes, 2, 
                        globalChapterSearchResult, pageChaptersSearchResult,
                        isInBrackets, isExcluded, isImportantVerse, force, !needToQueueIfChapter, processAsExtendedVerse, out localHierarchySearchResult, ref processedVerses, null);
                }

                first = false;

                System.Windows.Forms.Application.DoEvents();
            }           

            #endregion           

            #region linking rubbish notes pages

            if (SettingsManager.Instance.RubbishPage_Use)
            {
                var rubbishVerses = new List<VersePointer>();

                if (SettingsManager.Instance.RubbishPage_ExpandMultiVersesLinking && searchResult.VersePointer.IsMultiVerse)
                {
                    if (allSubVerses == null)
                        allSubVerses = searchResult.VersePointer.GetAllVerses(ref oneNoteApp,
                                                            new GetAllIncludedVersesExceptFirstArgs() { BibleNotebookId = SettingsManager.Instance.NotebookId_Bible });
                    rubbishVerses.AddRange(allSubVerses);
                }
                else
                    rubbishVerses.Add(searchResult.VersePointer);
                
                foreach (VersePointer vp in rubbishVerses)
                {                    
                    TryLinkVerseToNotesPage(ref oneNoteApp, vp, verseWeight, searchResult.ResultType, versePosition,
                        notePageId, notePageContentObjectId, linkDepth,
                        false, SettingsManager.Instance.RubbishPage_ExcludedVersesLinking, 
                        NotesPageType.Detailed, SettingsManager.Instance.PageWidth_RubbishNotes, 1, 
                        globalChapterSearchResult, pageChaptersSearchResult,
                        isInBrackets, isExcluded, isImportantVerse, force, !needToQueueIfChapter, processAsExtendedVerse, out localHierarchySearchResult, ref processedVerses, null);                

                    System.Windows.Forms.Application.DoEvents();
                }
            }

            #endregion

            return textElementValue;
        }

        private void CheckIfLinkIsChanged(VerseRecognitionManager.LinkInfo linkInfo, string textElementValue, int startVerseNameIndex, int endVerseNameIndex, string linkHref, string link)
        {
            if (linkInfo.IsLink)
            {
                var existingLinkHref = StringUtils.GetAttributeValue(textElementValue.Substring(startVerseNameIndex, endVerseNameIndex - startVerseNameIndex), "href");
                if (!string.IsNullOrEmpty(existingLinkHref))
                {
                    existingLinkHref = Uri.UnescapeDataString(existingLinkHref);
                    if (string.IsNullOrEmpty(linkHref))
                        linkHref = StringUtils.GetAttributeValue(link, "href");
                    if (existingLinkHref != linkHref)
                        _pageContentWasChanged = true;
                }
                else
                    _pageContentWasChanged = true;
            }
            else
                _pageContentWasChanged = true;
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
                    verseNumber.Value.IsMultiVerse ? BibleCommon.Resources.Constants.Verses : BibleCommon.Resources.Constants.Verse );

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
        /// <param name="excludedVersesLinking">Привязываем ли стихи, даже если они входят в главу, являющуюся ExcludableChapter</param>
        /// <param name="globalChapterSearchResult"></param>
        /// <param name="pageChaptersSearchResult"></param>
        /// <param name="isInBrackets"></param>
        /// <param name="force"></param>
        /// <param name="hierarchySearchResult"></param>
        /// <param name="onHierarchyElementFound"></param>
        /// <param name="notesPageLevel">1, 2 or 3</param>        
        /// <returns></returns>
        private bool TryLinkVerseToNotesPage(ref Application oneNoteApp, VersePointer vp, decimal verseWeight,
            VersePointerSearchResult.SearchResultType resultType, XmlCursorPosition versePosition,
            HierarchyElementInfo notePageId, string notePageContentObjectId, AnalyzeDepth linkDepth,
            bool createLinkToNotesPage, bool excludedVersesLinking, 
            NotesPageType notesPageType, int notesPageWidth, int notesPageLevel, 
            VersePointerSearchResult globalChapterSearchResult, List<VersePointerSearchResult> pageChaptersSearchResult,
            bool isInBrackets, bool isExcluded, bool isImportantVerse, bool force, bool forceAnalyzeChapter, bool processAsExtendedVerse,
            out BibleSearchResult hierarchySearchResult, ref List<SimpleVersePointer> processedVerses,
            Action<BibleSearchResult> onHierarchyElementFound)
        {
            hierarchySearchResult = BibleHierarchySearchProvider.GetHierarchyObject(ref oneNoteApp, ref vp, linkDepth);            

            if (hierarchySearchResult.FoundSuccessfully)
            {
                if (hierarchySearchResult.HierarchyObjectInfo.VerseNumber.HasValue)
                    vp.VerseNumber = hierarchySearchResult.HierarchyObjectInfo.VerseNumber.Value;    // то есть мы уточняем номер для этого стиха (ведь может быть смежные стихи, как в ibs). Важно: здесь у нас точно !vp.IsMultiVerse, так как если бы он был IsMultiVerse, то мы его переделали в subVerse (и у него теперь есть ParentVersePointer)

                if (hierarchySearchResult.HierarchyObjectInfo == null 
                    || hierarchySearchResult.HierarchyObjectInfo.PageId != notePageId.Id)
                {
                    if (onHierarchyElementFound != null)
                        onHierarchyElementFound(hierarchySearchResult);

                    //bool isChapter = VersePointerSearchResult.IsChapter(resultType);
                    var isChapter = vp.IsChapter;

                    //if ((isChapter ^ vp.IsChapter))  // данные расходятся. что-то тут не чисто. запишем варнинг и ничего делать не будем
                    //{
                    //    Logger.LogWarning("Invalid verse result: '{0}' - '{1}'", vp.OriginalVerseName, resultType);
                    //}
                    //else
                    //{
                        if ((!isChapter || excludedVersesLinking || forceAnalyzeChapter) && linkDepth >= AnalyzeDepth.Full)   // главы сразу не обрабатываем - вдруг есть стихи этих глав в текущей заметке. Вот если нет - тогда потом и обработаем. Но если у нас стоит excludedVersesLinking, то сразу обрабатываем
                        {
                            bool canContinue = true;

                            if (!excludedVersesLinking)  // иначе всё равно привязываем
                            {
                                if (isExcluded || IsExcludedCurrentNotePage)
                                    canContinue = false;

                                if (canContinue)
                                {                                    
                                    if (VersePointerSearchResult.IsVerseWithoutChapter(resultType))
                                    {
                                        if (globalChapterSearchResult != null)
                                        {
                                            if (globalChapterSearchResult.ResultType == VersePointerSearchResult.SearchResultType.ChapterAndVerseAtStartString)
                                            {
                                                if (!isInBrackets)
                                                {
                                                    if (globalChapterSearchResult.VersePointer.IsMultiVerse)
                                                        if (globalChapterSearchResult.VersePointer.IsInVerseRange(vp))
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
                                var verses = LinkVerseToNotesPage(ref oneNoteApp, vp, verseWeight, versePosition, isChapter,
                                    hierarchySearchResult.HierarchyObjectInfo,
                                    notePageId,
                                    notePageContentObjectId, createLinkToNotesPage, notesPageType, notesPageWidth,
                                    notesPageLevel, isImportantVerse, force, processAsExtendedVerse);

                                if (processedVerses != null)
                                    processedVerses.AddRange(verses);
                            }
                        }

                        return true;
                    //}
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
        private List<SimpleVersePointer> LinkVerseToNotesPage(ref Application oneNoteApp, VersePointer vp, decimal verseWeight, XmlCursorPosition versePosition, bool isChapter,
            BibleHierarchyObjectInfo verseHierarchyObjectInfo,
            HierarchyElementInfo notePageId, string notePageContentObjectId, bool createLinkToNotesPage,
            NotesPageType notesPageType, int notesPageWidth, int notesPageLevel, bool isImportantVerse, bool force, bool processAsExtendedVerse)
        {
            bool notesPageWasModified;
            var notesPageName = GetNotesPageName(notesPageType, vp, verseHierarchyObjectInfo);

            if (!SettingsManager.Instance.StoreNotesPagesInFolder)
            {
                notesPageWasModified = LinkVerseToNotesPageInOneNote(ref oneNoteApp, vp, verseWeight, versePosition, isChapter, verseHierarchyObjectInfo, notePageId, notePageContentObjectId, createLinkToNotesPage,
                    notesPageType, notesPageName, notesPageWidth, notesPageLevel, isImportantVerse, force, processAsExtendedVerse);
            }
            else
            {
                notesPageWasModified = LinkVerseToNotesPageInFileSystem(ref oneNoteApp, vp, verseWeight, versePosition, isChapter, verseHierarchyObjectInfo, notePageId, notePageContentObjectId, createLinkToNotesPage,
                    notesPageType, notesPageName, notesPageWidth, notesPageLevel, isImportantVerse, force, processAsExtendedVerse);
            }

            var key = new NotePageProcessedVerseId() { NotePageId = notePageId.UniqueName, NotesPageName = notesPageName };
            var processedVerses = AddNotePageProcessedVerse(key, vp, verseHierarchyObjectInfo.VerseNumber);            

            if (createLinkToNotesPage && (notesPageWasModified || force || isChapter))
            {
                OneNoteProxy.Instance.AddProcessedBiblePageWithUpdatedLinksToNotesPages(vp.GetChapterPointer(), verseHierarchyObjectInfo);

                OneNoteProxy.Instance.AddProcessedVerseOnBiblePageWithUpdatedLinksToNotesPages(processedVerses);  // добавляем только стихи, отмеченные на "Сводной заметок"
            }

            return processedVerses;
        }

        private bool LinkVerseToNotesPageInFileSystem(ref Application oneNoteApp, VersePointer vp, decimal verseWeight, XmlCursorPosition versePosition, bool isChapter,
            BibleHierarchyObjectInfo verseHierarchyObjectInfo,
            HierarchyElementInfo notePageId, string notePageContentObjectId, bool createLinkToNotesPage,
            NotesPageType notesPageType, string notesPageName, int notesPageWidth, int notesPageLevel, bool isImportantVerse, bool force, bool processAsExtendedVerse)
        {
            var pageWasAdded = NotesPageManagerFS.UpdateNotesPage(ref oneNoteApp, this, vp, verseWeight, versePosition, isChapter, 
                                        verseHierarchyObjectInfo,
                                        notePageId, notePageContentObjectId, notesPageType, notesPageName,
                                        isImportantVerse, force, processAsExtendedVerse, !(AnalyzeAllPages && force));

            return pageWasAdded;
        }


        public static string GetNotesPageName(NotesPageType notesPageType, VersePointer vp, BibleHierarchyObjectInfo verseHierarchyObjectInfo)
        {
            string notesPageName;
            switch (notesPageType)
            {
                case NotesPageType.Verse:
                    notesPageName = GetDefaultNotesPageName(
                                verseHierarchyObjectInfo.AdditionalObjectsIds.ContainsKey(vp)
                                    ? (VerseNumber?)verseHierarchyObjectInfo.AdditionalObjectsIds[vp].VerseNumber
                                    : verseHierarchyObjectInfo.VerseNumber);                    
                    break;
                case NotesPageType.Chapter:
                    notesPageName = SettingsManager.Instance.PageName_Notes;
                    break;
                case NotesPageType.Detailed:
                    notesPageName = SettingsManager.Instance.PageName_RubbishNotes;
                    break;
                default:
                    throw new NotSupportedException(notesPageType.ToString());
            }

            return notesPageName;
        }

        private bool LinkVerseToNotesPageInOneNote(ref Application oneNoteApp, VersePointer vp, decimal verseWeight, XmlCursorPosition versePosition, bool isChapter,
            BibleHierarchyObjectInfo verseHierarchyObjectInfo,
            HierarchyElementInfo notePageId, string notePageContentObjectId, bool createLinkToNotesPage,
            NotesPageType notesPageType, string notesPageName, int notesPageWidth, int notesPageLevel, bool isImportantVerse, bool force, bool processAsExtendedVerse)
        {
            string biblePageName = verseHierarchyObjectInfo.PageName;
            
            string notesParentPageName = notesPageType == NotesPageType.Verse ? SettingsManager.Instance.PageName_Notes : null;            

            string notesPageId = null;
            bool pageWasCreated = false;
            bool rowWasAdded = false;
            try
            {
                notesPageId = OneNoteProxy.Instance.GetNotesPageId(ref oneNoteApp,
                    verseHierarchyObjectInfo.SectionId,
                    verseHierarchyObjectInfo.PageId, biblePageName, notesPageName, out pageWasCreated, notesParentPageName, notesPageLevel);
            }
            catch (NotFoundVerseLinkPageExceptions)
            {   
                // возможно дело в устаревшем кэше                

                if (verseHierarchyObjectInfo.LoadedFromCache)
                {
                    var fullSearchResult = HierarchySearchManager.GetHierarchyObject(
                                                ref oneNoteApp, SettingsManager.Instance.NotebookId_Bible, ref vp, HierarchySearchManager.FindVerseLevel.AllVerses, false);

                    if (fullSearchResult.FoundSuccessfully)
                    {
                        if (fullSearchResult.HierarchyObjectInfo.PageId != notePageId.Id)
                        {
                            verseHierarchyObjectInfo = fullSearchResult.HierarchyObjectInfo;
                            try
                            {
                                notesPageId = OneNoteProxy.Instance.GetNotesPageId(ref oneNoteApp,
                                          verseHierarchyObjectInfo.SectionId,
                                          verseHierarchyObjectInfo.PageId, biblePageName, notesPageName, out pageWasCreated, notesParentPageName, notesPageLevel);

                                if (!string.IsNullOrEmpty(notesPageId))  // значит действительно кэш устарел. Надо его удалить и написать об этом предупреждение                                                   
                                {
                                    OneNoteProxy.Instance.CleanBibleVersesLinksCache();
                                    Logger.LogWarning(BibleCommon.Resources.Constants.BibleVersesLinksCacheWasCleaned);
                                }
                            }
                            catch (Exception ex)
                            {
                                Logger.LogError(ex);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogError(ex);
            }

            if (!string.IsNullOrEmpty(notesPageId))
            {
                string targetContentObjectId = _notesPagesProviderManager.UpdateNotesPage(ref oneNoteApp, this, vp, verseWeight, versePosition, isChapter, verseHierarchyObjectInfo,
                        notePageId, notesPageId, notePageContentObjectId, 
                        notesPageName, notesPageWidth, isImportantVerse, force, processAsExtendedVerse, out rowWasAdded);

                //if (createLinkToNotesPage && (pageWasCreated || rowWasAdded || force))
                //{
                //    OneNoteProxy.PageContent versePageDocument = OneNoteProxy.Instance.GetPageContent(ref oneNoteApp, verseHierarchyObjectInfo.PageId, OneNoteProxy.PageType.Bible);       

                //    string link = string.Format("<font size='2pt'>{0}</font>",
                //                    OneNoteUtils.GenerateLink(ref oneNoteApp, SettingsManager.Instance.PageName_Notes, notesPageId, null)); // здесь всегда передаём null, так как в частых случаях он и так null, потому что страница в кэше, и в OneNote она ещё не обновлялась (то есть идентификаторы ещё не проставлены). Так как эти идентификаторы проставятся в самом конце, то и ссылки обновим в конце.

                //    bool wasModified = false;

                //    if (isChapter)
                //    {
                //        if (SetLinkToNotesPageForChapter(versePageDocument.Content, link, versePageDocument.Xnm))
                //            wasModified = true;
                //    }
                //    else
                //    {
                //        if (SetLinkToNotesPageForVerse(versePageDocument.Content, link, vp, verseHierarchyObjectInfo, versePageDocument.Xnm))
                //            wasModified = true;
                //    }

                //    if (wasModified)
                //    {
                //        versePageDocument.WasModified = true;
                        
                //        OneNoteProxy.Instance.AddProcessedBiblePageWithUpdatedLinksToNotesPages(verseHierarchyObjectInfo.SectionId, verseHierarchyObjectInfo.PageId, biblePageName, vp.GetChapterPointer());

                //        OneNoteProxy.Instance.AddProcessedVerseOnBiblePageWithUpdatedLinksToNotesPages(vp, verseHierarchyObjectInfo.VerseNumber);  // добавляем только стихи, отмеченные на "Сводной заметок"
                //    }                    
                //}

                var keyForOldProvider = new NotePageProcessedVerseId() { NotePageId = notePageId.Id, NotesPageName = notesPageName };
                AddNotePageProcessedVerseForOldProvider(keyForOldProvider, vp, verseHierarchyObjectInfo.VerseNumber);

                //var key = new NotePageProcessedVerseId() { NotePageId = notePageId.ManualId ?? notePageId.Id, NotesPageName = notesPageName };
                //return AddNotePageProcessedVerse(key, vp, verseHierarchyObjectInfo.VerseNumber);

                return pageWasCreated || rowWasAdded;
            }

            return false;
        }




        //private bool SetLinkToNotesPageForVerse(XDocument pageDocument, string link, VersePointer vp,
        //    HierarchySearchManager.HierarchyObjectInfo verseHierarchyObjectInfo, XmlNamespaceManager xnm)
        //{
        //    bool result = false;

        //    //находим ячейку для заметок стиха
        //    XElement contentObject = pageDocument.XPathSelectElement(string.Format("//one:OE[@objectID = \"{0}\"]",
        //        verseHierarchyObjectInfo.VerseContentObjectId), xnm);
        //    if (contentObject == null)
        //    {
        //        Logger.LogError("{0} '{1}'", BibleCommon.Resources.Constants.NoteLinkManagerVerseCellNotFound,  vp.OriginalVerseName);
        //    }
        //    else
        //    {
        //        XNode cellForNotesPage = contentObject.Parent.Parent.NextNode;
        //        XElement textElement = cellForNotesPage.XPathSelectElement("one:OEChildren/one:OE/one:T", xnm);

        //        //if (textElement.Value == string.Empty)  // лучше обновлять ссылку на страницу заметок, так как зачастую она вначале бывает неточной (только на страницу)
        //        {
        //            textElement.Value = link;
        //            result = true;
        //        }
        //    }

        //    return result;
        //}


        /// <summary>
        /// Возвращает элемент - ссылку на страницу сводную заметок на ряд для главы
        /// </summary>
        /// <param name="pageDocument"></param>
        /// <param name="xnm"></param>
        /// <returns></returns>
        public static XElement GetChapterNotesPageLink(XDocument pageDocument, XmlNamespaceManager xnm)
        {
            XElement notesLinkElement = null;

            var oeElMeta = pageDocument.Root.XPathSelectElement(string.Format("one:Outline/one:OEChildren/one:OE/one:Meta[@name=\"{0}\"]", Constants.Key_NotesPageLink), xnm);
            if (oeElMeta != null)
            {
                notesLinkElement = oeElMeta.Parent.XPathSelectElement("one:T", xnm);
            }
            else  // пробуем искать по-старому (для обратной совместимости)
            {
                notesLinkElement = pageDocument.Root.XPathSelectElement("one:Outline/one:OEChildren", xnm);

                if (notesLinkElement != null && notesLinkElement.Nodes().Count() == 1)   // похоже на правду                        
                {
                    notesLinkElement = notesLinkElement.XPathSelectElement(string.Format("one:OE/one:T[contains(.,'>{0}<')]",
                        SettingsManager.Instance.PageName_Notes), xnm);

                    if (notesLinkElement != null)
                    {
                        OneNoteUtils.UpdateElementMetaData(notesLinkElement.Parent, Constants.Key_NotesPageLink, "1", xnm);
                        UpdateChapterNotesPageLinkPosition(
                                notesLinkElement, 
                                SettingsManager.Instance.PageWidth_Bible + Constants.ChapterNotesPageLinkOutline_OffsetX,
                                Constants.ChapterNotesPageLinkOutline_y,
                                xnm);
                    }
                }
                else 
                    notesLinkElement = null;
            }

            return notesLinkElement;
        }

        public static void UpdateChapterNotesPageLinkPosition(XElement chapterNotesPageLinkEl, int x, int y, XmlNamespaceManager xnm)
        {
            if (chapterNotesPageLinkEl != null)
            {
                var positionEl = chapterNotesPageLinkEl.Parent.Parent.Parent.XPathSelectElement("one:Position", xnm);
                if (positionEl != null)
                {
                    positionEl.SetAttributeValue("x", x);
                    positionEl.SetAttributeValue("y", y);
                }
            }
        }

        public static XElement GetChapterNotesPageLinkAndCreateIfNeeded(XDocument pageDocument, XmlNamespaceManager xnm)
        {   
            var notesLinkElement = GetChapterNotesPageLink(pageDocument, xnm);                

            if (notesLinkElement == null)
            {
                XElement titleElement = pageDocument.Root.XPathSelectElement("one:Title", xnm);

                XNamespace nms = XNamespace.Get(Constants.OneNoteXmlNs);
                var outlineEl = new XElement(nms + "Outline",
                                 new XElement(nms + "Position",
                                        new XAttribute("x", SettingsManager.Instance.PageWidth_Bible + Constants.ChapterNotesPageLinkOutline_OffsetX),
                                        new XAttribute("y", Constants.ChapterNotesPageLinkOutline_y),
                                        new XAttribute("z", Constants.ChapterNotesPageLinkOutline_z)));
                var oeChildrenEl = new XElement(nms + "OEChildren");

                notesLinkElement = new XElement(nms + "T",
                            new XCData(string.Empty));

                var oeEl = new XElement(nms + "OE", 
                            notesLinkElement);

                OneNoteUtils.UpdateElementMetaData(oeEl, Constants.Key_NotesPageLink, "1", xnm);

                oeChildrenEl.Add(oeEl);
                outlineEl.Add(oeChildrenEl);
                titleElement.AddAfterSelf(outlineEl);                
            }

            return notesLinkElement;
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

        public List<SimpleVersePointer> AddNotePageProcessedVerseForOldProvider(NotePageProcessedVerseId verseId, VersePointer vp, VerseNumber? verseNumber)
        {
            if (!_notePageProcessedVersesForOldProvider.ContainsKey(verseId))
            {
                _notePageProcessedVersesForOldProvider.Add(verseId, new HashSet<SimpleVersePointer>());
            }

            var svp = vp.ToSimpleVersePointer();
            if (verseNumber.HasValue)
                svp.VerseNumber = verseNumber.Value;

            var result = svp.GetAllVerses();

            if (!_notePageProcessedVersesForOldProvider[verseId].Contains(svp))   // отслеживаем обработанные стихи для каждой из страниц сводной заметок
            {
                result.ForEach(v => _notePageProcessedVersesForOldProvider[verseId].Add(v));
            }

            return result;
        }

        public bool ContainsNotePageProcessedVerseForOldProvider(NotePageProcessedVerseId verseId, VersePointer vp)
        {
            if (!_notePageProcessedVersesForOldProvider.ContainsKey(verseId))
            {
                _notePageProcessedVersesForOldProvider.Add(verseId, new HashSet<SimpleVersePointer>());
            }

            return _notePageProcessedVersesForOldProvider[verseId].Contains(vp.ToSimpleVersePointer());
        }
      
        #endregion


        public void Dispose()
        {
            _oneNoteApp = null;
        }
    }
}
