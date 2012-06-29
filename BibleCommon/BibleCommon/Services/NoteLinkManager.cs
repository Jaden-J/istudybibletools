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
            
            public ProcessVerseEventArgs(bool foundVerse)
            {
                FoundVerse = foundVerse;
            }            
        }

      

        private class FoundChapterInfo  // в итоге здесь будут только те главы, которые представлены в текущей заметке без стихов
        {
            public string TextElementObjectId { get; set; }
            public VersePointerSearchResult VersePointerSearchResult { get; set; }
            public HierarchySearchManager.HierarchySearchResult HierarchySearchResult { get; set; }
        }    

        #endregion


        internal const char ChapterVerseDelimiter = ':';
        //internal static readonly char[] SymbolsAfterBibleVerse = new string[] { };
        internal const string DoNotAnalyzeAllPageSymbol = "{}";       
          

        public enum AnalyzeDepth
        {
            OnlyFindVerses = 1,            
            GetVersesLinks = 2,
            Full = 3
        }

       

        internal bool IsExcludedCurrentNotePage { get; set; }
        private Dictionary<NotePageProcessedVerseId, HashSet<VersePointer>> _notePageProcessedVerses = new Dictionary<NotePageProcessedVerseId, HashSet<VersePointer>>();  

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
                    if (linkDepth > AnalyzeDepth.GetVersesLinks)
                        linkDepth = AnalyzeDepth.GetVersesLinks;  // на странице заметок только обновляем ссылки
                    notePageDocument.PageType = OneNoteProxy.PageType.NotesPage;  // уточняем тип страницы
                }

                if (notePageName.Contains(DoNotAnalyzeAllPageSymbol))
                    IsExcludedCurrentNotePage = true;

                XElement titleElement = notePageDocument.Content.Root.XPathSelectElement("one:Title/one:OE", notePageDocument.Xnm);
                string pageTitleId = titleElement != null ? titleElement.Attribute("objectID").Value : null;


                string noteSectionGroupName = OneNoteUtils.GetHierarchyElementName(_oneNoteApp, sectionGroupId);
                string noteSectionName = OneNoteUtils.GetHierarchyElementName(_oneNoteApp, sectionId);
                List<FoundChapterInfo> foundChapters = new List<FoundChapterInfo>();
                List<VersePointerSearchResult> pageChaptersSearchResult = ProcessPageTitle(_oneNoteApp, notePageDocument.Content,
                    noteSectionGroupName, noteSectionName, notePageName, pageId, pageTitleId, foundChapters, notePageDocument.Xnm, linkDepth, force, isSummaryNotesPage,
                    out wasModified);  // получаем главы текущей страницы, указанные в заголовке (глобальные главы, если больше одной - то не используем их при определении принадлежности только ситхов (:3))                

                List<XElement> processedTextElements = new List<XElement>();

                foreach (XElement oeChildrenElement in notePageDocument.Content.Root.XPathSelectElements("one:Outline/one:OEChildren", notePageDocument.Xnm))
                {
                    if (ProcessTextElements(_oneNoteApp, oeChildrenElement, noteSectionGroupName, noteSectionName,
                         notePageName, pageId, pageTitleId, foundChapters, processedTextElements, pageChaptersSearchResult,
                         notePageDocument.Xnm, linkDepth, force, isSummaryNotesPage))
                        wasModified = true;
                }

                if (foundChapters.Count > 0)  // то есть имеются главы, которые указаны в тексте именно как главы, без стихов, и на которые надо делать тоже ссылки
                {
                    ProcessChapters(foundChapters, noteSectionGroupName, noteSectionName, notePageName, pageId, pageTitleId, linkDepth, force);                                       
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

        private void ProcessChapters(List<FoundChapterInfo> foundChapters, 
            string noteSectionGroupName, string noteSectionName, string notePageName, string pageId, string pageTitleId, 
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
                            noteSectionGroupName, noteSectionName, notePageName, pageId, pageTitleId, chapterInfo.TextElementObjectId, true,
                            SettingsManager.Instance.PageName_Notes, null, SettingsManager.Instance.PageWidth_Notes, 1,
                            chapterInfo.VersePointerSearchResult.ResultType == VersePointerSearchResult.SearchResultType.ExcludableChapter ? true : force);
                    }

                    if (SettingsManager.Instance.RubbishPage_Use)
                    {
                        if (!SettingsManager.Instance.RubbishPage_ExcludedVersesLinking)   // иначе мы её обработали сразу же, когда встретили
                        {
                            LinkVerseToNotesPage(_oneNoteApp, chapterInfo.VersePointerSearchResult.VersePointer, true,
                                chapterInfo.HierarchySearchResult.HierarchyObjectInfo,
                                noteSectionGroupName, noteSectionName, notePageName, pageId, pageTitleId, chapterInfo.TextElementObjectId, false,
                                SettingsManager.Instance.PageName_RubbishNotes, null, SettingsManager.Instance.PageWidth_RubbishNotes, 1,
                                chapterInfo.VersePointerSearchResult.ResultType == VersePointerSearchResult.SearchResultType.ExcludableChapter ? true : force);
                        }
                    }
                }
            }

            Logger.LogMessage(string.Empty, false, true, false);
        }

        private List<VersePointerSearchResult> ProcessPageTitle(Application oneNoteApp, XDocument notePageDocument,
            string noteSectionGroupName, string noteSectionName, string notePageName, string pageId, string pageTitleId,
            List<FoundChapterInfo> foundChapters, XmlNamespaceManager xnm,
            AnalyzeDepth linkDepth, bool force, bool isSummaryNotesPage, out bool wasModified)
        {
            wasModified = false;
            List<VersePointerSearchResult> pageChaptersSearchResult = new List<VersePointerSearchResult>();
            VersePointerSearchResult globalChapterSearchResult = null;   // результат поиска "глобальной" главы 
            VersePointerSearchResult prevResult = null;

            if (ProcessTextElement(oneNoteApp, notePageDocument.Root.XPathSelectElement("one:Title/one:OE/one:T", xnm),
                        noteSectionGroupName, noteSectionName, notePageName, pageId, pageTitleId,
                        foundChapters, ref globalChapterSearchResult, ref prevResult, null, linkDepth, force, true, isSummaryNotesPage, searchResult =>
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
            string noteSectionGroupName, string noteSectionName, string notePageName, string pageId, string pageTitleId,
            List<FoundChapterInfo> foundChapters, List<XElement> processedTextElements,
            List<VersePointerSearchResult> pageChaptersSearchResult,
            XmlNamespaceManager xnm, AnalyzeDepth linkDepth, bool force, bool isSummaryNotesPage)
        {
            bool wasModified = false;

            VersePointerSearchResult globalChapterSearchResult;   // результат поиска "глобальной" главы 
            VersePointerSearchResult prevResult;            

            foreach (XElement cellElement in parent.XPathSelectElements(".//one:Table/one:Row/one:Cell", xnm))
            {
                if (ProcessTextElements(oneNoteApp, cellElement, noteSectionGroupName, noteSectionName,
                        notePageName, pageId, pageTitleId, foundChapters, processedTextElements, pageChaptersSearchResult,
                        xnm, linkDepth, force, isSummaryNotesPage))
                    wasModified = true;
            }
            
            globalChapterSearchResult = pageChaptersSearchResult.Count == 1 ? pageChaptersSearchResult[0] : null;   // если в заголовке указана одна глава - то используем её при нахождении только стихов, если же указано несколько - то не используем их
            prevResult = null;            

            foreach (XElement textElement in parent.XPathSelectElements(".//one:T", xnm))
            {
                if (processedTextElements.Contains(textElement))
                {
                    globalChapterSearchResult = pageChaptersSearchResult.Count == 1 ? pageChaptersSearchResult[0] : null;   // если в заголовке указана одна глава - то используем её при нахождении только стихов, если же указано несколько - то не используем их
                    prevResult = null;                    
                    continue;
                }

                if (ProcessTextElement(oneNoteApp, textElement, noteSectionGroupName, noteSectionName,
                                         notePageName, pageId, pageTitleId, foundChapters,
                                         ref globalChapterSearchResult, ref prevResult, pageChaptersSearchResult, linkDepth, force, false, isSummaryNotesPage, null))
                    wasModified = true;

                processedTextElements.Add(textElement);
            }


            return wasModified;
        }

        private void FireProcessVerseEvent(bool foundVerse)
        {
            if (OnNextVerseProcess != null)
            {
                var args = new ProcessVerseEventArgs(foundVerse);
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
        /// <param name="pageId"></param>
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
        private bool ProcessTextElement(Application oneNoteApp, XElement textElement, string noteSectionGroupName,
            string noteSectionName, string notePageName, string pageId, string pageTitleId, List<FoundChapterInfo> foundChapters,
            ref VersePointerSearchResult globalChapterSearchResult, ref VersePointerSearchResult prevResult, 
            List<VersePointerSearchResult> pageChaptersSearchResult,
            AnalyzeDepth linkDepth, bool force, bool isTitle, bool isSummaryNotesPage, Action<VersePointerSearchResult> onVersePointerFound)
        {

            FireProcessVerseEvent(false);

            bool wasModified = false;
            string localChapterName = string.Empty;    // имя главы в пределах данного стиха. например, действительно только для девятки в "Откр 5:7,9"

            if (textElement != null && !string.IsNullOrEmpty(textElement.Value))
            {
                OneNoteUtils.NormalizaTextElement(textElement);
                string textElementValue = textElement.Value;
                int numberIndex = textElement.Value
                        .IndexOfAny(new char[] { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9' });
                
                while (numberIndex > -1)
                {
                    try
                    {
                        int number;
                        int textBreakIndex;
                        int htmlBreakIndex;
                        bool isLink;
                        bool isInBrackets;
                        bool isExcluded;
                        if (CanProcessAtNumberPosition(textElement, numberIndex, 
                            out number, out textBreakIndex, out htmlBreakIndex, out isLink, out isInBrackets, out isExcluded))
                        {
                            VersePointerSearchResult searchResult = GetValidVersePointer(textElement,
                                numberIndex, textBreakIndex - 1, number,
                                globalChapterSearchResult,
                                localChapterName, prevResult, isInBrackets, isTitle);

                            if (searchResult.ResultType != VersePointerSearchResult.SearchResultType.Nothing && isSummaryNotesPage)
                                if (searchResult.VersePointer != null && searchResult.VersePointer.IsMultiVerse)  // если находимся на странице сводной заметок и нашли мультивёрс ссылку (например :4-7) - то такие ссылки не обрабатываем
                                {
                                    searchResult.ResultType = VersePointerSearchResult.SearchResultType.Nothing;
                                    numberIndex = searchResult.VersePointerHtmlEndIndex;
                                }

                            if (searchResult.ResultType != VersePointerSearchResult.SearchResultType.Nothing)
                            {
                                FireProcessVerseEvent(true);

                                if (onVersePointerFound != null)
                                    onVersePointerFound(searchResult);                                

                                string textToChange;
                                HierarchySearchManager.HierarchySearchResult hierarchySearchResult;

                                if (searchResult.ResultType == VersePointerSearchResult.SearchResultType.ChapterOnlyAtStartString
                                    || searchResult.ResultType == VersePointerSearchResult.SearchResultType.ChapterAndVerseAtStartString)      // считаем, что указан глава только тогда, когда конкретно указана только глава, а не глава со стихом, например
                                {
                                    globalChapterSearchResult = searchResult;
                                }

                                localChapterName = searchResult.ChapterName;   // всегда запоминаем


                                if (!isLink || (isLink && force) || (isTitle && isInBrackets))
                                {
                                    if (VersePointerSearchResult.IsVerse(searchResult.ResultType))
                                        textToChange = searchResult.VerseString;
                                    else if (VersePointerSearchResult.IsChapter(searchResult.ResultType))
                                        textToChange = searchResult.ChapterName;
                                    else
                                        textToChange = searchResult.VersePointer.OriginalVerseName;

                                    string textElementObjectId = (string)textElement.Parent.Attribute("objectID");


                                    textElementValue = ProcessVerse(oneNoteApp, searchResult,                                                                
                                                                textToChange,
                                                                textElementValue,
                                                                noteSectionGroupName, noteSectionName,
                                                                notePageName, pageId, pageTitleId, textElementObjectId,
                                                                linkDepth, globalChapterSearchResult, pageChaptersSearchResult,
                                                                isLink, isInBrackets, isExcluded, force, out numberIndex, out hierarchySearchResult);

                                    if (searchResult.ResultType == VersePointerSearchResult.SearchResultType.SingleVerseOnly)  // то есть нашли стих, а до этого значит была скорее всего просто глава!
                                    {
                                        FoundChapterInfo chapterInfo = foundChapters.FirstOrDefault(fch =>
                                                fch.VersePointerSearchResult.ResultType != VersePointerSearchResult.SearchResultType.ExcludableChapter
                                                 && fch.VersePointerSearchResult.VersePointer.ChapterName == searchResult.VersePointer.ChapterName);

                                        if (chapterInfo != null)
                                            foundChapters.Remove(chapterInfo);
                                    }
                                    else if (VersePointerSearchResult.IsChapter(searchResult.ResultType))
                                    {
                                        if (hierarchySearchResult.ResultType == HierarchySearchManager.HierarchySearchResultType.Successfully
                                            && hierarchySearchResult.HierarchyStage == HierarchySearchManager.HierarchyStage.Page)
                                        {
                                            if (!foundChapters.Exists(fch =>
                                                    fch.VersePointerSearchResult.ResultType == VersePointerSearchResult.SearchResultType.ExcludableChapter
                                                     && fch.VersePointerSearchResult.VersePointer.ChapterName == searchResult.VersePointer.ChapterName))
                                            {
                                                foundChapters.Add(new FoundChapterInfo()
                                                {
                                                    TextElementObjectId = textElementObjectId,
                                                    VersePointerSearchResult = searchResult,
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
                                    numberIndex = searchResult.VersePointerHtmlEndIndex;

                                prevResult = searchResult;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.LogError(ex);
                    }

                    if (textElement.Value.Length >= numberIndex + 2)
                         numberIndex = textElement.Value
                              .IndexOfAny(new char[] { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9' }, numberIndex + 1);
                    else
                        numberIndex = -1;
                }
            }

            return wasModified;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="textElement"></param>
        /// <param name="numberIndex"></param>
        /// <param name="number"></param>
        /// <param name="breakIndex"></param>
        /// <param name="isLink"></param>
        /// <param name="isInBrackets">если в скобках []</param>
        /// <param name="isExcluded">Не стоит его по дефолту анализировать</param>
        /// <returns></returns>
        private bool CanProcessAtNumberPosition(XElement textElement, int numberIndex, 
            out int number, out int textBreakIndex, out int htmlBreakIndex, out bool isLink, out bool isInBrackets, out bool isExcluded)
        {
            isLink = false;
            number = -1;
            textBreakIndex = -1;
            htmlBreakIndex = -1;
            isInBrackets = false;
            isExcluded = false;

            if (numberIndex == 0)  // не может начинаться с цифры
                return false;

            char prevChar = StringUtils.GetChar(textElement.Value, numberIndex - 1);
            string numberString = StringUtils.GetNextString(textElement.Value, numberIndex - 1, null, out textBreakIndex, out htmlBreakIndex);
            char nextChar = StringUtils.GetChar(textElement.Value, htmlBreakIndex);

            if (((prevChar == '>' || prevChar == ' ' || prevChar == '.' || StringUtils.IsCharAlphabetical(prevChar))                             // нашли полную ссылку
                    && (nextChar == '<' || nextChar == ChapterVerseDelimiter 
                            || nextChar == default(char) || nextChar == ' '
                            || nextChar == ')' || nextChar == ']' || nextChar == ',' || nextChar == '.' 
                            || nextChar == '?' || nextChar == '!') || nextChar == '-')  // нашли ссылку только на главу
                ||
                ((prevChar == ChapterVerseDelimiter || prevChar == '>' || prevChar == ',')                                                // нашли только стих
                    && !StringUtils.IsCharAlphabetical(nextChar)))
            {
                number = int.Parse(numberString);  // полюбому тут должно быть число 

                if (number > 0 && number <= 176)
                {
                    isLink = StringUtils.IsSurroundedBy(textElement.Value, "<a", "</a", numberIndex);
                    isInBrackets = StringUtils.IsSurroundedBy(textElement.Value, "[", "]", numberIndex);
                    isExcluded = StringUtils.IsSurroundedBy(textElement.Value, "{", "}", numberIndex);      

                    return true;
                }
            }
            
            return false;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="textElement"></param>
        /// <param name="numberStartIndex"></param>
        /// <param name="numberEndIndex"></param>
        /// <param name="number"></param>
        /// <param name="globalChapterName"></param>
        /// <param name="localChapterName">только в пределах строки</param>
        /// <param name="prevResult"></param>
        /// <returns></returns>
        private VersePointerSearchResult GetValidVersePointer(XElement textElement,
            int numberStartIndex, int numberEndIndex, int number, VersePointerSearchResult globalChapterResult, string localChapterName,
            VersePointerSearchResult prevResult, bool isInBrackets, bool isTitle)
        {
            VersePointerSearchResult result = new VersePointerSearchResult() 
                                    { ResultType = VersePointerSearchResult.SearchResultType.Nothing };
            int startIndex;
            int endIndex;
            int nextTextBreakIndex;
            int nextHtmlBreakIndex;
            int prevTextBreakIndex;
            int prevHtmlBreakIndex;            

            string prevChar = StringUtils.GetPrevString(textElement.Value, numberStartIndex, null, out prevTextBreakIndex, out prevHtmlBreakIndex, StringSearchIgnorance.None, StringSearchMode.SearchFirstValueChar);
            string nextChar = StringUtils.GetNextString(textElement.Value, numberEndIndex, null, out nextTextBreakIndex, out nextHtmlBreakIndex, StringSearchIgnorance.None, StringSearchMode.SearchFirstValueChar);

            if (nextChar == ChapterVerseDelimiter.ToString())  // как будто нашли полную ссылку
            {
                string verseString = StringUtils.GetNextString(textElement.Value, nextHtmlBreakIndex, null, out endIndex, out nextHtmlBreakIndex);
                int verseNumber;
                if (int.TryParse(verseString, out verseNumber))
                {
                    if (verseNumber <= 176 && verseNumber > 0)
                    {
                        verseString = GetFullVerseString(textElement.Value, verseString, ref endIndex, ref nextHtmlBreakIndex);                      

                        for (int maxMissCount = 2; maxMissCount >= 0; maxMissCount--)
                        {
                            string bookName = StringUtils.GetPrevString(textElement.Value,
                                numberStartIndex, new SearchMissInfo(maxMissCount, SearchMissInfo.MissMode.CancelOnMissFound), out startIndex, out prevHtmlBreakIndex,
                                maxMissCount > 0 ? StringSearchIgnorance.IgnoreAllSpacesAndDots : StringSearchIgnorance.IgnoreFirstSpaceAndDot,
                                StringSearchMode.SearchText);

                            if (!string.IsNullOrEmpty(bookName) && !string.IsNullOrEmpty(bookName.Trim()))
                            {
                                char prevPrevChar = StringUtils.GetChar(textElement.Value, prevHtmlBreakIndex);
                                if (!(StringUtils.IsCharAlphabetical(prevPrevChar) || StringUtils.IsDigit(prevPrevChar)))
                                {
                                    string verseName = string.Format("{0}{1}{2}{3}{4}", bookName, bookName.EndsWith(" ") ? string.Empty : " ", number, ChapterVerseDelimiter, verseString);

                                    VersePointer vp = new VersePointer(verseName);
                                    if (vp.IsValid)
                                    {
                                        result.VersePointer = vp;
                                        result.ResultType = prevHtmlBreakIndex == -1
                                                    ? VersePointerSearchResult.SearchResultType.ChapterAndVerseAtStartString
                                                    : VersePointerSearchResult.SearchResultType.ChapterAndVerse;
                                        result.VersePointerEndIndex = endIndex;
                                        result.VersePointerStartIndex = startIndex + 1;
                                        result.VersePointerHtmlEndIndex = nextHtmlBreakIndex;
                                        result.VersePointerHtmlStartIndex = prevHtmlBreakIndex;
                                        result.ChapterName = string.Format("{0}{1}{2}", bookName, bookName.EndsWith(" ") ? string.Empty : " ", number);
                                        result.TextElement = textElement;
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            else if (prevChar == ChapterVerseDelimiter.ToString() || prevChar == ",") // как будто нашли ссылку только на стих                    
            {
                int temp, temp2;
                string prevPrevChar = StringUtils.GetPrevString(textElement.Value, prevHtmlBreakIndex, null, out temp, out temp2, StringSearchIgnorance.None, StringSearchMode.SearchFirstChar);                
                string globalChapterName = globalChapterResult != null ? globalChapterResult.ChapterName : string.Empty;

                if (!string.IsNullOrEmpty(globalChapterName) || !string.IsNullOrEmpty(localChapterName))
                {
                    bool canContinue = true;

                    string chapterName = globalChapterName;

                    VersePointerSearchResult.SearchResultType resultType = VersePointerSearchResult.SearchResultType.SingleVerseOnly;

                    if (prevChar == ",")    // надо проверить, чтоб предыдущий символ тоже был цифрой   
                    {
                        chapterName = localChapterName;
                        resultType = VersePointerSearchResult.SearchResultType.FollowingVerseOnly;

                        if (prevResult == null)
                            canContinue = false;
                        else
                        {
                            if (!int.TryParse(prevPrevChar, out temp))
                            {
                                canContinue = false;
                            }
                            else
                            {
                                if (!(VersePointerSearchResult.IsChapterAndVerse(prevResult.ResultType)
                                       || VersePointerSearchResult.IsVerse(prevResult.ResultType)))  // если только что до этого мы нашли стих, тогда запятая - это разделитель стихов, иначе здесь может иметься ввиду разделение глав, которое пока не поддерживаем
                                {
                                    canContinue = false;
                                }
                            }

                            if (canContinue)
                            {
                                if (prevHtmlBreakIndex < prevResult.VersePointerHtmlEndIndex)
                                    canContinue = false;
                                else
                                {
                                    string stringBetweenThisAndLastResultHTML = textElement.Value
                                                                    .Substring(prevResult.VersePointerHtmlEndIndex, prevHtmlBreakIndex - prevResult.VersePointerHtmlEndIndex + 1);
                                    if (StringUtils.GetText(stringBetweenThisAndLastResultHTML).Length > 1)  // то есть не рядом был предыдущий результат
                                        canContinue = false;
                                }
                            }
                        }
                    }
                    else if (prevChar == ChapterVerseDelimiter.ToString())    // надо проверить, чтоб предыдущий символ не был цифрой и не буквой
                    {                        
                        if (int.TryParse(prevPrevChar, out temp))
                            canContinue = false;

                        if (canContinue && !string.IsNullOrEmpty(prevPrevChar))
                        {
                            if (StringUtils.IsCharAlphabetical(prevPrevChar[0]))
                                canContinue = false;
                        }

                        if (canContinue)
                        {
                            if (globalChapterResult == null)
                                canContinue = false;
                            else
                                if (globalChapterResult.DepthLevel > OneNoteUtils.GetDepthLevel(textElement))  // то есть глава находится глубже (правее), чем найденный стих. должно быть либо на одном уровне, либо левее.
                                    canContinue = false;
                            
                        }
                    }

                    if (canContinue)
                    {
                        if (!string.IsNullOrEmpty(chapterName))
                        {
                            string verseString = number.ToString();
                            endIndex = nextTextBreakIndex;

                            verseString = GetFullVerseString(textElement.Value, verseString, ref endIndex, ref nextHtmlBreakIndex);

                            string verseName = string.Format("{0}{1}{2}", chapterName, ChapterVerseDelimiter, verseString);

                            VersePointer vp = new VersePointer(verseName);
                            if (vp.IsValid)
                            {
                                result.VersePointer = vp;
                                result.ResultType = resultType;
                                result.VersePointerEndIndex = endIndex;
                                result.VersePointerStartIndex = prevChar == ChapterVerseDelimiter.ToString() ? prevTextBreakIndex: prevTextBreakIndex + 1;
                                result.VersePointerHtmlEndIndex = nextHtmlBreakIndex;
                                result.VersePointerHtmlStartIndex = prevHtmlBreakIndex;
                                result.VerseString = (prevChar == ChapterVerseDelimiter.ToString() ? prevChar : string.Empty) + verseString;
                                result.ChapterName = chapterName;
                                result.TextElement = textElement;
                            }
                        }
                    }
                }
            }
            else if (string.IsNullOrEmpty(nextChar.Trim()) || nextChar == ")" || nextChar == "]" || nextChar == "," || nextChar == "." || nextChar == "?" || nextChar == "!")  // как будто нашли только ссылку на главу
            {
                for (int maxMissCount = 2; maxMissCount >= 0; maxMissCount--)
                {
                    string bookName = StringUtils.GetPrevString(textElement.Value,
                        numberStartIndex, new SearchMissInfo(maxMissCount, SearchMissInfo.MissMode.CancelOnMissFound), out startIndex, out prevHtmlBreakIndex,
                        maxMissCount > 0 ? StringSearchIgnorance.IgnoreAllSpacesAndDots : StringSearchIgnorance.IgnoreFirstSpaceAndDot,
                        StringSearchMode.SearchText);                    

                    if (!string.IsNullOrEmpty(bookName) && !string.IsNullOrEmpty(bookName.Trim()))
                    {
                        if (isInBrackets)
                        {
                            if (textElement.Value[startIndex + 1] == '[')  // временное решение проблемы: когда указана в заголовке только глава в квадратных скобках, то первая скобка удаляется 
                                startIndex++;                           
                            
                            bookName = bookName.Trim('[', ']');
                        }

                        char prevPrevChar = StringUtils.GetChar(textElement.Value, prevHtmlBreakIndex);
                        if (!(StringUtils.IsCharAlphabetical(prevPrevChar) || StringUtils.IsDigit(prevPrevChar)))
                        {
                            string verseName = string.Format("{0}{1}{2}{3}{4}", bookName, bookName.EndsWith(" ") ? string.Empty : " ", number, ChapterVerseDelimiter, 0);

                            VersePointer vp = new VersePointer(verseName);
                            if (vp.IsValid)
                            {
                                result.ChapterName = string.Format("{0}{1}{2}", bookName, bookName.EndsWith(" ") ? string.Empty : " ", number);
                                bool chapterOnlyAtStartString = prevHtmlBreakIndex == -1 && nextChar != "-";    // так как пока мы не поддерживаем Откр 1-2, и так как мы найдём в таком случае только Откр 1, то не считаем это ссылкой, как ChapterOnlyAtStartString
                                if (isInBrackets && isTitle)
                                    result.ResultType = VersePointerSearchResult.SearchResultType.ExcludableChapter;
                                else
                                    result.ResultType = chapterOnlyAtStartString                                     
                                                ? VersePointerSearchResult.SearchResultType.ChapterOnlyAtStartString
                                                : VersePointerSearchResult.SearchResultType.ChapterOnly;
                                result.VersePointerEndIndex = nextTextBreakIndex;
                                result.VersePointerStartIndex = startIndex + 1;
                                result.VersePointerHtmlEndIndex = nextHtmlBreakIndex;
                                result.VersePointerHtmlStartIndex = prevHtmlBreakIndex;
                                result.VersePointer = vp;
                                result.TextElement = textElement;
                                break;
                            }
                        }
                    }
                }
            }

            return result;
        }

        private static string GetFullVerseString(string textElementValue, string verseString, ref int endIndex, ref int nextHtmlBreakIndex)
        {
            if (StringUtils.GetChar(textElementValue, nextHtmlBreakIndex) == '-')
            {
                char tempNextChar = StringUtils.GetChar(textElementValue, nextHtmlBreakIndex + 1);
                if (StringUtils.IsDigit(tempNextChar))
                {
                    verseString = string.Format("{0}-{1}",
                        verseString,
                        StringUtils.GetNextString(textElementValue, nextHtmlBreakIndex, null, out endIndex, out nextHtmlBreakIndex));  // чтоб учесть случай Откр 4:5-9 - чтоб определить, где заканчивается ссылка
                }
            }

            return verseString;
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
        /// <param name="isLink"></param>        
        /// <param name="force"></param>        
        /// <param name="newEndVerseIndex"></param>
        /// <param name="hierarchySearchResult"></param>
        /// <returns></returns>
        private string ProcessVerse(Application oneNoteApp, VersePointerSearchResult searchResult,
            string textToChange, string textElementValue, string noteSectionGroupName, string noteSectionName, 
            string notePageName, string notePageId, string notePageTitleId, string notePageContentObjectId,
            AnalyzeDepth linkDepth, VersePointerSearchResult globalChapterSearchResult, List<VersePointerSearchResult> pageChaptersSearchResult,
            bool isLink, bool isInBrackets, bool isExcluded, bool force, out int newEndVerseIndex, out HierarchySearchManager.HierarchySearchResult hierarchySearchResult)
        {
            hierarchySearchResult = new HierarchySearchManager.HierarchySearchResult() { ResultType = HierarchySearchManager.HierarchySearchResultType.NotFound };

            int startVerseNameIndex = searchResult.VersePointerStartIndex;
            int endVerseNameIndex = searchResult.VersePointerEndIndex;

            newEndVerseIndex = endVerseNameIndex;

            if (!CorrectTextToChangeBoundary(textElementValue, isLink,
                              ref startVerseNameIndex, ref endVerseNameIndex))
            {
                newEndVerseIndex = searchResult.VersePointerHtmlEndIndex; // потому что это же значение мы присваиваем, если стоит !force и встретили гиперссылку                
                return textElementValue;
            }


            #region linking main notes page            

            List<VersePointer> verses = new List<VersePointer>() { searchResult.VersePointer };

            if (SettingsManager.Instance.ExpandMultiVersesLinking && searchResult.VersePointer.IsMultiVerse)
                verses.AddRange(searchResult.VersePointer.GetAllIncludedVersesExceptFirst());

            bool first = true;
            foreach (VersePointer vp in verses)
            {
                if (TryLinkVerseToNotesPage(oneNoteApp, vp, searchResult.ResultType,
                        noteSectionGroupName, noteSectionName, notePageName, notePageId, notePageTitleId, notePageContentObjectId, linkDepth,
                        !SettingsManager.Instance.UseDifferentPagesForEachVerse, SettingsManager.Instance.ExcludedVersesLinking,
                        SettingsManager.Instance.PageName_Notes, null, SettingsManager.Instance.PageWidth_Notes, 1,
                        globalChapterSearchResult, pageChaptersSearchResult,
                        isInBrackets, isExcluded, force, out hierarchySearchResult, hsr =>
                        {
                            if (first)
                            {
                                if (hsr.HierarchyStage == HierarchySearchManager.HierarchyStage.ContentPlaceholder)
                                    Logger.LogMessage("{0}: {1}", BibleCommon.Resources.Constants.ProcessVerse, searchResult.VersePointer.OriginalVerseName);
                                else
                                    Logger.LogMessage("{0}: {1}", BibleCommon.Resources.Constants.ProcessChapter, textToChange);
                            }
                        }))
                {
                    if (first)
                    {
                        if (linkDepth >= AnalyzeDepth.GetVersesLinks)
                        {
                            string link = OneNoteUtils.GenerateHref(oneNoteApp, textToChange,
                                hierarchySearchResult.HierarchyObjectInfo.PageId, hierarchySearchResult.HierarchyObjectInfo.ContentObjectId);

                            link = string.Format("<span style='font-weight:normal'>{0}</span>", link);

                            textElementValue = string.Concat(
                                textElementValue.Substring(0, startVerseNameIndex),
                                link,
                                textElementValue.Substring(endVerseNameIndex));

                            newEndVerseIndex = startVerseNameIndex + link.Length;
                            searchResult.VersePointerHtmlEndIndex = newEndVerseIndex;
                        }
                    }
                }

                if (SettingsManager.Instance.UseDifferentPagesForEachVerse)  // для каждого стиха своя страница
                {
                    string notesPageName = GetDefaultNotesPageName(vp.Verse);
                    TryLinkVerseToNotesPage(oneNoteApp, vp, searchResult.ResultType,
                        noteSectionGroupName, noteSectionName, notePageName, notePageId, notePageTitleId, notePageContentObjectId, linkDepth,
                        true, SettingsManager.Instance.ExcludedVersesLinking,
                        notesPageName, SettingsManager.Instance.PageName_Notes, SettingsManager.Instance.PageWidth_Notes, 2,
                        globalChapterSearchResult, pageChaptersSearchResult,
                        isInBrackets, isExcluded, force, out hierarchySearchResult, null);
                }

                first = false;
            }           

            #endregion           

            #region linking rubbish notes pages

            if (SettingsManager.Instance.RubbishPage_Use)
            {
                List<VersePointer> rubbishVerses = new List<VersePointer>() { searchResult.VersePointer };

                if (SettingsManager.Instance.RubbishPage_ExpandMultiVersesLinking && searchResult.VersePointer.IsMultiVerse)
                    rubbishVerses.AddRange(searchResult.VersePointer.GetAllIncludedVersesExceptFirst());

                foreach (VersePointer vp in rubbishVerses)
                {
                    TryLinkVerseToNotesPage(oneNoteApp, vp, searchResult.ResultType, 
                        noteSectionGroupName, noteSectionName, notePageName, notePageId, notePageTitleId, notePageContentObjectId, linkDepth,
                        false, SettingsManager.Instance.RubbishPage_ExcludedVersesLinking, 
                        SettingsManager.Instance.PageName_RubbishNotes, null, SettingsManager.Instance.PageWidth_RubbishNotes, 1,
                        globalChapterSearchResult, pageChaptersSearchResult,
                        isInBrackets, isExcluded, force, out hierarchySearchResult, null);
                }
            }

            #endregion

            return textElementValue;
        }

        internal static string GetDefaultNotesPageName(int? verseNumber)
        {
            if (verseNumber.GetValueOrDefault(0) > 0 && SettingsManager.Instance.UseDifferentPagesForEachVerse)
                return string.Format("{1} {2}", SettingsManager.Instance.PageName_Notes, verseNumber, BibleCommon.Resources.Constants.Verse);

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
            string noteSectionGroupName, string noteSectionName, string notePageName, string notePageId,
            string notePageTitleId, string notePageContentObjectId, AnalyzeDepth linkDepth,
            bool createLinkToNotesPage, bool excludedVersesLinking, 
            string notesPageName, string notesParentPageName, int notesPageWidth, int notesPageLevel,
            VersePointerSearchResult globalChapterSearchResult, List<VersePointerSearchResult> pageChaptersSearchResult,
            bool isInBrackets, bool isExcluded, bool force, out HierarchySearchManager.HierarchySearchResult hierarchySearchResult,
            Action<HierarchySearchManager.HierarchySearchResult> onHierarchyElementFound)
        {

            hierarchySearchResult = HierarchySearchManager.GetHierarchyObject(
                                                oneNoteApp, SettingsManager.Instance.NotebookId_Bible, vp);
            if (hierarchySearchResult.ResultType == HierarchySearchManager.HierarchySearchResultType.Successfully)
            {
                if (hierarchySearchResult.HierarchyStage == HierarchySearchManager.HierarchyStage.ContentPlaceholder
                    || hierarchySearchResult.HierarchyStage == HierarchySearchManager.HierarchyStage.Page)
                {
                    if (hierarchySearchResult.HierarchyObjectInfo.PageId != notePageId)
                    {
                        if (onHierarchyElementFound != null)
                            onHierarchyElementFound(hierarchySearchResult);

                        bool isChapter = VersePointerSearchResult.IsChapter(resultType);

                        if ((!isChapter || excludedVersesLinking) && linkDepth >= AnalyzeDepth.Full)   // главы сразу не обрабатываем - вдруг есть стихи этих глав в текущей заметке. Вот если нет - тогда потом и обработаем. Но если у нас стоит excludedVersesLinking, то сразу обрабатываем
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
                                                return pcsr.ResultType == VersePointerSearchResult.SearchResultType.ExcludableChapter
                                                        && pcsr.VersePointer.ChapterName == vp.ChapterName;
                                            }))
                                                canContinue = false;  // то есть среди исключаемых глав есть текущая
                                        }
                                    }
                                }
                            }

                            if (canContinue)
                            {
                                LinkVerseToNotesPage(oneNoteApp, vp, isChapter,
                                    hierarchySearchResult.HierarchyObjectInfo,
                                    noteSectionGroupName, noteSectionName, notePageName, notePageId, notePageTitleId,
                                    notePageContentObjectId, createLinkToNotesPage, notesPageName, notesParentPageName, notesPageWidth, notesPageLevel, force);
                            }
                        }

                        return true;
                    }
                }
            }

            return false;
        }


        
        /// <summary>
        /// 
        /// </summary>
        /// <param name="textElementValue"></param>
        /// <param name="isLink"></param>
        /// <param name="startVerseNameIndex"></param>
        /// <param name="endVerseNameIndex"></param>
        /// <returns>false - если помимо библейской ссылки, в гиперссылке содержится и другой текст. Не обрабатываем такие ссылки</returns>
        private bool CorrectTextToChangeBoundary(string textElementValue, bool isLink, ref int startVerseNameIndex, ref int endVerseNameIndex)
        {
            if (isLink)
            {
                string beginSearchString = "<a ";
                string endSearchString = "</a>";
                int linkStartIndex = StringUtils.LastIndexOf(textElementValue, beginSearchString, 0, startVerseNameIndex);
                if (linkStartIndex != -1)
                {
                    int linkEndIndex = textElementValue.IndexOf(endSearchString, endVerseNameIndex);
                    if (linkEndIndex != -1)
                    {
                        int startVerseNameIndexTemp = linkStartIndex;
                        int endVerseNameIndexTemp = linkEndIndex + endSearchString.Length;

                        string textBefore = StringUtils.GetText(textElementValue.Substring(startVerseNameIndexTemp, startVerseNameIndex - startVerseNameIndexTemp));                        
                        if (string.IsNullOrEmpty(textBefore))                        
                        {
                            string textAfter = StringUtils.GetText(textElementValue.Substring(endVerseNameIndex, endVerseNameIndexTemp - endVerseNameIndex));
                            if (string.IsNullOrEmpty(textAfter))
                            {
                                startVerseNameIndex = startVerseNameIndexTemp;
                                endVerseNameIndex = endVerseNameIndexTemp;
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
        private void LinkVerseToNotesPage(Application oneNoteApp, VersePointer vp, bool isChapter,
            HierarchySearchManager.HierarchyObjectInfo verseHierarchyObjectInfo,
            string noteSectionGroupName, string noteSectionName, string notePageName, string notePageId, string notePageTitleId, string notePageContentObjectId, bool createLinkToNotesPage,
            string notesPageName, string notesParentPageName, int notesPageWidth, int notesPageLevel, bool force)
        {            
            OneNoteProxy.PageContent versePageDocument = OneNoteProxy.Instance.GetPageContent(oneNoteApp, verseHierarchyObjectInfo.PageId, OneNoteProxy.PageType.Bible);       
            string biblePageName = (string)versePageDocument.Content.Root.Attribute("name");

            string notesPageId = null;
            try
            {
                notesPageId = OneNoteProxy.Instance.GetNotesPageId(oneNoteApp,
                    verseHierarchyObjectInfo.SectionId,
                    verseHierarchyObjectInfo.PageId, biblePageName, notesPageName, notesParentPageName, notesPageLevel);                    
            }
            catch (Exception ex)
            {
                Logger.LogError(ex);
            }

            if (!string.IsNullOrEmpty(notesPageId))
            {
                string targetContentObjectId = UpdateNotesPage(oneNoteApp, vp, isChapter, verseHierarchyObjectInfo,
                        noteSectionGroupName, noteSectionName, notesPageId, notePageName, notePageId, notePageTitleId, notePageContentObjectId, 
                        notesPageName, notesPageWidth, force);

                if (createLinkToNotesPage)
                {
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
                        OneNoteProxy.Instance.AddProcessedBiblePages(verseHierarchyObjectInfo.SectionId, verseHierarchyObjectInfo.PageId, biblePageName, vp.GetChapterPointer());

                        OneNoteProxy.Instance.AddProcessedVerse(vp);  // добавляем только стихи, отмеченные на "Сводной заметок"
                    }                    
                }

                var key = new NotePageProcessedVerseId() { NotePageId = notePageId, NotesPageName = notesPageName };
                AddNotePageProcessedVerse(key, vp);                
            }
        }

        private bool SetLinkToNotesPageForVerse(XDocument pageDocument, string link, VersePointer vp,
            HierarchySearchManager.HierarchyObjectInfo verseHierarchyObjectInfo, XmlNamespaceManager xnm)
        {
            bool result = false;

            //находим ячейку для заметок стиха
            XElement contentObject = pageDocument.XPathSelectElement(string.Format("//one:OE[@objectID = '{0}']",
                verseHierarchyObjectInfo.ContentObjectId), xnm);
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

        private string UpdateNotesPage(Application oneNoteApp, VersePointer vp, bool isChapter,
            HierarchySearchManager.HierarchyObjectInfo verseHierarchyObjectInfo,
            string noteSectionGroupName, string noteSectionName,
            string notesPageId, string notePageName, string notePageId, string notePageTitleId, string notePageContentObjectId,
            string notesPageName, int notesPageWidth, bool force)
        {
            string targetContentObjectId = string.Empty;            
            XNamespace nms = XNamespace.Get(Constants.OneNoteXmlNs);
            OneNoteProxy.PageContent notesPageDocument = OneNoteProxy.Instance.GetPageContent(oneNoteApp, notesPageId, OneNoteProxy.PageType.NotesPage);

            XElement rowElement = GetNotesRowAndCreateIfNotExists(oneNoteApp, vp, isChapter, notesPageWidth, verseHierarchyObjectInfo,
                notesPageDocument.Content, notesPageDocument.Xnm, nms);            

            if (rowElement != null)
            {
                AddLinkToNotePage(oneNoteApp, vp, rowElement, noteSectionGroupName, noteSectionName,
                    notePageName, notePageId, notePageTitleId, notePageContentObjectId, notesPageDocument, notesPageDocument.Xnm, nms, notesPageName, force);

                targetContentObjectId = GetNotesRowObjectId(oneNoteApp, notesPageId, vp.Verse, isChapter);               
            }            

            return targetContentObjectId;
        }

        private void AddLinkToNotePage(Application oneNoteApp, VersePointer vp, XElement rowElement, 
            string noteSectionGroupName, string noteSectionName,
            string notePageName, string notePageId, string notePageTitleId, string notePageContentObjectId,
            OneNoteProxy.PageContent notesPageDocument, XmlNamespaceManager xnm, XNamespace nms, string notesPageName, bool force)
        {
            string noteTitle = (noteSectionGroupName != noteSectionName && !string.IsNullOrEmpty(noteSectionGroupName))
                ? string.Format("{0} / {1} / {2}", noteSectionGroupName, noteSectionName, notePageName)
                : string.Format("{0} / {1}", noteSectionName, notePageName);
            
            XElement suchNoteLink = null;            
            XElement notesCellElement = rowElement.XPathSelectElement("one:Cell[2]/one:OEChildren", xnm);

            string link = OneNoteUtils.GenerateHref(oneNoteApp, noteTitle, notePageId, notePageContentObjectId);
            int pageIdStringIndex = link.IndexOf("page-id={");
            if (pageIdStringIndex != -1)
            {
                string pageId = link.Substring(pageIdStringIndex, link.IndexOf('}', pageIdStringIndex) - pageIdStringIndex + 1);                
                suchNoteLink = rowElement.XPathSelectElement(string.Format(
                   "one:Cell[2]/one:OEChildren/one:OE/one:T[contains(.,'{0}')]", pageId), xnm);                
            }

            if (suchNoteLink != null)
            {
                var key = new NotePageProcessedVerseId() { NotePageId = notePageId, NotesPageName = notesPageName };
                if (force && !ContainsNotePageProcessedVerse(key, vp))  // если в первый раз и force                
                {  // удаляем старые ссылки на текущую странцу, так как мы начали новый анализ с параметром "force" и мы только в первый раз зашли сюда
                    var verseLinks = suchNoteLink.Parent.NextNode;
                    if (verseLinks != null && verseLinks.XPathSelectElement("one:List", xnm) == null)
                        verseLinks.Remove();

                    suchNoteLink.Parent.Remove();
                    suchNoteLink = null;
                }
            }

            if (suchNoteLink != null)
                OneNoteUtils.NormalizaTextElement(suchNoteLink);         

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
                string pageLink = OneNoteUtils.GenerateHref(oneNoteApp, noteTitle, notePageId, notePageTitleId);                

                var verseLinksOE = suchNoteLink.Parent.NextNode;
                if (verseLinksOE != null && verseLinksOE.XPathSelectElement("one:List", xnm) == null)  // значит следующая строка без номера, то есть значит идут ссылки
                {
                    XElement existingVerseLinksElement = verseLinksOE.XPathSelectElement("one:T", xnm);


                    int currentVerseIndex = existingVerseLinksElement.Value.Split(new string[] { "</a>" }, StringSplitOptions.None).Length;

                    existingVerseLinksElement.Value += Resources.Constants.VerseLinksDelimiter + OneNoteUtils.GenerateHref(oneNoteApp,
                                string.Format(Resources.Constants.VerseLinkTemplate, currentVerseIndex), notePageId, notePageContentObjectId)
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
                                                                    string.Format(Resources.Constants.VerseLinkTemplate, 2), notePageId, notePageContentObjectId)
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
                return string.Format(" <b>(:{0}-{1})</b>", vp.Verse, vp.TopVerse);
            else
                return string.Empty;
        }

        private string GetExistingMultiVerseString(XElement suchNoteLink)
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

        private XElement GetNotesRowAndCreateIfNotExists(Application oneNoteApp, VersePointer vp, bool isChapter, int mainColumnWidth, HierarchySearchManager.HierarchyObjectInfo verseHierarchyObjectInfo, 
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

        private void AddNewNotesRow(Application oneNoteApp, VersePointer vp, bool isChapter, HierarchySearchManager.HierarchyObjectInfo verseHierarchyObjectInfo,
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

        #region helper methods

        public void AddNotePageProcessedVerse(NotePageProcessedVerseId verseId, VersePointer vp)
        {
            if (!_notePageProcessedVerses.ContainsKey(verseId))
            {
                _notePageProcessedVerses.Add(verseId, new HashSet<VersePointer>());
            }

            if (!_notePageProcessedVerses[verseId].Contains(vp))   // отслеживаем обработанные стихи для каждой из страниц сводной заметок
            {
                _notePageProcessedVerses[verseId].Add(vp);
            }
        }

        public bool ContainsNotePageProcessedVerse(NotePageProcessedVerseId verseId, VersePointer vp)
        {
            if (!_notePageProcessedVerses.ContainsKey(verseId))
            {
                _notePageProcessedVerses.Add(verseId, new HashSet<VersePointer>());
            }

            return _notePageProcessedVerses[verseId].Contains(vp);
        }

        #endregion


        public void Dispose()
        {
            _oneNoteApp = null;
        }
    }
}
