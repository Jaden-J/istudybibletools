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

namespace BibleNoteLinker
{    
    public static class NoteLinkManager
    {        
        internal const char ChapterVerseDelimiter = ':';
        
        private class FoundChapterInfo  // в итоге здесь будут только те главы, которые представлены в текущей заметке без стихов
        {
            public string TextElementObjectId { get; set; }
            public VersePointerSearchResult VersePointerSearchResult { get; set; }
            public HierarchySearchManager.HierarchySearchResult HierarchySearchResult { get; set; }            
        }        

        public enum AnalyzeDepth
        {
            OnlyFindVerses = 1,            
            GetVersesLinks = 2,
            Full = 3
        }        

        /// <summary>
        /// 
        /// </summary>
        /// <param name="oneNoteApp"></param>
        /// <param name="sectionGroupId"></param>
        /// <param name="sectionId"></param>
        /// <param name="pageId"></param>
        /// <param name="linkDepth"></param>
        /// <param name="force">Обрабатывать даже ссылки</param>
        public static void LinkPageVerses(Application oneNoteApp, string sectionGroupId, string sectionId, string pageId, 
            AnalyzeDepth linkDepth, bool force)
        {
            try
            {
                bool wasModified = false;

                string pageContentXml;
                XDocument notePageDocument;
                XmlNamespaceManager xnm;
                oneNoteApp.GetPageContent(pageId, out pageContentXml);
                notePageDocument = OneNoteUtils.GetXDocument(pageContentXml, out xnm);
                string notePageName = (string)notePageDocument.Root.Attribute("name");

                if (!CanProcessPage(notePageDocument, notePageName))
                    return;

                string noteSectionGroupName = OneNoteUtils.GetHierarchyElementName(oneNoteApp, sectionGroupId);
                string noteSectionName = OneNoteUtils.GetHierarchyElementName(oneNoteApp, sectionId);
                List<FoundChapterInfo> foundChapters = new List<FoundChapterInfo>();
                List<VersePointer> processedVerses = new List<VersePointer>();   // отслеживаем обработанные стихи, чтобы, например, верно подсчитывать verseCount, когда анализируем ссылки в том числе
                List<VersePointerSearchResult> pageChaptersSearchResult = ProcessPageTitle(oneNoteApp, notePageDocument,
                    noteSectionGroupName, noteSectionName, notePageName, pageId, processedVerses, foundChapters, xnm, linkDepth, force,
                    out wasModified);  // получаем главы текущей страницы, указанные в заголовке (глобальные главы, если больше одной - то не используем их при определении принадлежности только ситхов (:3))                

                List<XElement> processedTextElements = new List<XElement>();

                foreach (XElement oeChildrenElement in notePageDocument.Root.XPathSelectElements("one:Outline/one:OEChildren", xnm))
                {
                    if (ProcessTextElements(oneNoteApp, oeChildrenElement, noteSectionGroupName, noteSectionName,
                         notePageName, pageId, processedVerses, foundChapters, processedTextElements, pageChaptersSearchResult,
                         xnm, linkDepth, force))
                        wasModified = true;
                }

                if (foundChapters.Count > 0)  // то есть имеются главы, которые указаны в тексте именно как главы, без стихов, и на которые надо делать тоже ссылки
                {
                    Logger.LogMessage("Заключительная обработка глав", true, false);

                    foreach (FoundChapterInfo chapterInfo in foundChapters)
                    {
                        if (linkDepth >= AnalyzeDepth.Full)
                        {
                            Logger.LogMessage(".", false, false, false);
                            LinkVerseToNotesPage(oneNoteApp, chapterInfo.VersePointerSearchResult.VersePointer, true,
                                processedVerses, chapterInfo.HierarchySearchResult.HierarchyObjectInfo,
                                noteSectionGroupName, noteSectionName, notePageName, pageId, chapterInfo.TextElementObjectId,
                                chapterInfo.VersePointerSearchResult.ResultType == VersePointerSearchResult.SearchResultType.ExcludableChapter ? true : force);
                        }
                    }

                    Logger.LogMessage(string.Empty, false, true, false);
                }

                if (wasModified)
                    oneNoteApp.UpdatePageContent(notePageDocument.ToString());
            }
            catch (Exception ex)
            {
                Logger.LogError("Ошибки при обработке страницы.", ex);
            }
        }

        private static List<VersePointerSearchResult> ProcessPageTitle(Application oneNoteApp, XDocument notePageDocument,
            string noteSectionGroupName, string noteSectionName, string notePageName, string pageId,
            List<VersePointer> processedVerses, List<FoundChapterInfo> foundChapters, XmlNamespaceManager xnm,
            AnalyzeDepth linkDepth, bool force, out bool wasModified)
        {
            wasModified = false;
            List<VersePointerSearchResult> pageChaptersSearchResult = new List<VersePointerSearchResult>();
            VersePointerSearchResult globalChapterSearchResult = null;   // результат поиска "глобальной" главы 
            VersePointerSearchResult prevResult = null;

            if (ProcessTextElement(oneNoteApp, notePageDocument.Root.XPathSelectElement("one:Title/one:OE/one:T", xnm),
                        noteSectionGroupName, noteSectionName, notePageName, pageId,
                        processedVerses, foundChapters, ref globalChapterSearchResult, ref prevResult, null, linkDepth, force, true, searchResult =>
                        {
                            if (VersePointerSearchResult.IsChapter(searchResult.ResultType))
                                pageChaptersSearchResult.Add(searchResult);

                        }))
                wasModified = true;

            return pageChaptersSearchResult;
        }

        private static bool CanProcessPage(XDocument pageDocument, string pageName)
        {
            if (pageName.StartsWith(SettingsManager.Instance.PageName_Notes + "."))
                return false;

            return true;
        }

        private static bool ProcessTextElements(Application oneNoteApp, XElement parent,
            string noteSectionGroupName, string noteSectionName, string notePageName, string pageId,
            List<VersePointer> processedVerses, List<FoundChapterInfo> foundChapters, List<XElement> processedTextElements,
            List<VersePointerSearchResult> pageChaptersSearchResult,
            XmlNamespaceManager xnm, AnalyzeDepth linkDepth, bool force)
        {
            bool wasModified = false;

            VersePointerSearchResult globalChapterSearchResult;   // результат поиска "глобальной" главы 
            VersePointerSearchResult prevResult;            

            foreach (XElement cellElement in parent.XPathSelectElements(".//one:Table/one:Row/one:Cell", xnm))
            {
                if (ProcessTextElements(oneNoteApp, cellElement, noteSectionGroupName, noteSectionName,
                        notePageName, pageId, processedVerses, foundChapters, processedTextElements, pageChaptersSearchResult,
                        xnm, linkDepth, force))
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
                                         notePageName, pageId, processedVerses, foundChapters,
                                         ref globalChapterSearchResult, ref prevResult, pageChaptersSearchResult, linkDepth, force, false, null))
                    wasModified = true;

                processedTextElements.Add(textElement);
            }


            return wasModified;
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
        private static bool ProcessTextElement(Application oneNoteApp, XElement textElement, string noteSectionGroupName,
            string noteSectionName, string notePageName, string pageId, List<VersePointer> processedVerses, List<FoundChapterInfo> foundChapters,
            ref VersePointerSearchResult globalChapterSearchResult, ref VersePointerSearchResult prevResult, 
            List<VersePointerSearchResult> pageChaptersSearchResult,
            AnalyzeDepth linkDepth, bool force, bool isTitle, Action<VersePointerSearchResult> onVersePointerFound)
        {
            bool wasModified = false;
            string localChapterName = string.Empty;    // имя главы в пределах данного стиха. например, действительно только для девятки в "Откр 5:7,9"

            if (textElement != null && !string.IsNullOrEmpty(textElement.Value))
            {
                string textElementValue = textElement.Value;
                int numberIndex = textElement.Value
                        .IndexOfAny(new char[] { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9' });
                
                while (numberIndex > -1)
                {
                    try
                    {
                        int number;
                        int breakIndex;
                        bool isLink;
                        bool isInBrackets;
                        if (CanProcessAtNumberPosition(textElement, numberIndex, out number, out breakIndex, out isLink, out isInBrackets))
                        {
                            VersePointerSearchResult searchResult = GetValidVersePointer(textElement,
                                numberIndex, breakIndex - 1, number,
                                globalChapterSearchResult,
                                localChapterName, prevResult, isInBrackets, isTitle);

                            if (searchResult.ResultType != VersePointerSearchResult.SearchResultType.Nothing)
                            {
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


                                if (!isLink || (isLink && force))
                                {
                                    if (VersePointerSearchResult.IsVerse(searchResult.ResultType))
                                        textToChange = searchResult.VerseString;
                                    else if (VersePointerSearchResult.IsChapter(searchResult.ResultType))
                                        textToChange = searchResult.ChapterName;
                                    else
                                        textToChange = searchResult.VersePointer.OriginalVerseName;

                                    string textElementObjectId = (string)textElement.Parent.Attribute("objectID");

                                    textElementValue = ProcessVerse(oneNoteApp, searchResult,
                                                                processedVerses,
                                                                textToChange,
                                                                textElementValue,
                                                                noteSectionGroupName, noteSectionName,
                                                                notePageName, pageId, textElementObjectId,
                                                                linkDepth, globalChapterSearchResult, pageChaptersSearchResult,
                                                                isLink, isInBrackets, force, out numberIndex, out hierarchySearchResult);

                                    if (searchResult.ResultType == VersePointerSearchResult.SearchResultType.SingleVerseOnly)  // то есть нашли стих, а до этого значит была скорее всего просто глава!
                                    {
                                        List<FoundChapterInfo> chaptersInfo = foundChapters.Where(fch =>
                                                fch.VersePointerSearchResult.ResultType != VersePointerSearchResult.SearchResultType.ExcludableChapter
                                                 && fch.VersePointerSearchResult.VersePointer.ChapterName == searchResult.VersePointer.ChapterName).ToList();

                                        if (chaptersInfo.Count > 0)
                                            foundChapters.Remove(chaptersInfo[chaptersInfo.Count - 1]);
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
        /// <returns></returns>
        private static bool CanProcessAtNumberPosition(XElement textElement, int numberIndex, 
            out int number, out int breakIndex, out bool isLink, out bool isInBrackets)
        {
            isLink = false;
            number = -1;
            breakIndex = -1;
            isInBrackets = false;

            if (numberIndex == 0)  // не может начинаться с цифры
                return false;

            char prevChar = StringUtils.GetChar(textElement.Value, numberIndex - 1);
            string numberString = StringUtils.GetNextString(textElement.Value, numberIndex - 1, null, out breakIndex);
            char nextChar = StringUtils.GetChar(textElement.Value, breakIndex);

            if (((prevChar == '>' || prevChar == ' ' || prevChar == '.' || StringUtils.IsCharAlphabetical(prevChar))                             // нашли полную ссылку
                    && (nextChar == '<' || nextChar == ChapterVerseDelimiter 
                            || nextChar == default(char) || nextChar == ' '
                            || nextChar == ')' || nextChar == ']' || nextChar == ',' || nextChar == '.') || nextChar == '-')  // нашли ссылку только на главу
                ||
                ((prevChar == ChapterVerseDelimiter || prevChar == '>' || prevChar == ',')                                                // нашли только стих
                    && !StringUtils.IsCharAlphabetical(nextChar)))
            {
                number = int.Parse(numberString);  // полюбому тут должно быть число 

                if (number > 0 && number <= 176)
                {
                    isLink = StringUtils.IsSurroundedBy(textElement.Value, "<a", "</a", numberIndex);
                    isInBrackets = StringUtils.IsSurroundedBy(textElement.Value, "[", "]", numberIndex);                    

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
        private static VersePointerSearchResult GetValidVersePointer(XElement textElement,
            int numberStartIndex, int numberEndIndex, int number, VersePointerSearchResult globalChapterResult, string localChapterName,
            VersePointerSearchResult prevResult, bool isInBrackets, bool isTitle)
        {
            VersePointerSearchResult result = new VersePointerSearchResult() 
                                    { ResultType = VersePointerSearchResult.SearchResultType.Nothing };
            int startIndex;
            int endIndex;
            int nextBreakIndex;
            int prevBreakIndex;

            string prevChar = StringUtils.GetPrevString(textElement.Value, numberStartIndex, null, out prevBreakIndex, StringSearchIgnorance.None, StringSearchMode.SearchFirstValueChar);
            string nextChar = StringUtils.GetNextString(textElement.Value, numberEndIndex, null, out nextBreakIndex, StringSearchIgnorance.None, StringSearchMode.SearchFirstValueChar);

            if (nextChar == ChapterVerseDelimiter.ToString())  // как будто нашли полную ссылку
            {
                string verseString = StringUtils.GetNextString(textElement.Value, nextBreakIndex, null, out endIndex);
                int verseNumber;
                if (int.TryParse(verseString, out verseNumber))
                {
                    if (verseNumber <= 176 && verseNumber > 0)
                    {
                        verseString = GetFullVerseString(textElement.Value, verseString, ref endIndex);                      

                        for (int maxMissCount = 2; maxMissCount >= 0; maxMissCount--)
                        {
                            string bookName = StringUtils.GetPrevString(textElement.Value,
                                numberStartIndex, new SearchMissInfo(maxMissCount, SearchMissInfo.MissMode.CancelOnMissFound), out startIndex,
                                maxMissCount > 0 ? StringSearchIgnorance.IgnoreAllSpacesAndDots : StringSearchIgnorance.IgnoreFirstSpaceAndDot,
                                StringSearchMode.SearchText);

                            if (!string.IsNullOrEmpty(bookName) && !string.IsNullOrEmpty(bookName.Trim()))
                            {
                                char prevPrevChar = StringUtils.GetChar(textElement.Value, startIndex);
                                if (!(StringUtils.IsCharAlphabetical(prevPrevChar) || StringUtils.IsDigit(prevPrevChar)))
                                {
                                    string verseName = string.Format("{0}{1}{2}{3}{4}", bookName, bookName.EndsWith(" ") ? string.Empty : " ", number, ChapterVerseDelimiter, verseString);

                                    VersePointer vp = new VersePointer(verseName);
                                    if (vp.IsValid)
                                    {
                                        result.VersePointer = vp;
                                        result.ResultType = startIndex == -1
                                                    ? VersePointerSearchResult.SearchResultType.ChapterAndVerseAtStartString
                                                    : VersePointerSearchResult.SearchResultType.ChapterAndVerse;
                                        result.VersePointerEndIndex = endIndex;
                                        result.VersePointerStartIndex = startIndex;
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
                int temp;
                string prevPrevChar = StringUtils.GetPrevString(textElement.Value, prevBreakIndex, null, out temp, StringSearchIgnorance.None, StringSearchMode.SearchFirstChar);                
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
                                if (prevBreakIndex - prevResult.VersePointerEndIndex > 1) // то есть рядом был предыдущий результат
                                    canContinue = false;
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
                            endIndex = nextBreakIndex;

                            verseString = GetFullVerseString(textElement.Value, verseString, ref endIndex);

                            string verseName = string.Format("{0}{1}{2}", chapterName, ChapterVerseDelimiter, verseString);

                            VersePointer vp = new VersePointer(verseName);
                            if (vp.IsValid)
                            {
                                result.VersePointer = vp;
                                result.ResultType = resultType;
                                result.VersePointerEndIndex = endIndex;
                                result.VersePointerStartIndex = prevChar == ChapterVerseDelimiter.ToString() ? prevBreakIndex - 1 : prevBreakIndex;
                                result.VerseString = (prevChar == ChapterVerseDelimiter.ToString() ? prevChar : string.Empty) + verseString;
                                result.ChapterName = chapterName;
                                result.TextElement = textElement;
                            }
                        }
                    }
                }
            }
            else if (string.IsNullOrEmpty(nextChar.Trim()) || nextChar == ")" || nextChar == "]" || nextChar == "," || nextChar == ".")  // как будто нашли только ссылку на главу
            {
                for (int maxMissCount = 2; maxMissCount >= 0; maxMissCount--)
                {
                    string bookName = StringUtils.GetPrevString(textElement.Value,
                        numberStartIndex, new SearchMissInfo(maxMissCount, SearchMissInfo.MissMode.CancelOnMissFound), out startIndex,
                        maxMissCount > 0 ? StringSearchIgnorance.IgnoreAllSpacesAndDots : StringSearchIgnorance.IgnoreFirstSpaceAndDot,
                        StringSearchMode.SearchText);

                    if (!string.IsNullOrEmpty(bookName) && !string.IsNullOrEmpty(bookName.Trim()))
                    {
                        char prevPrevChar = StringUtils.GetChar(textElement.Value, startIndex);
                        if (!(StringUtils.IsCharAlphabetical(prevPrevChar) || StringUtils.IsDigit(prevPrevChar)))
                        {
                            string verseName = string.Format("{0}{1}{2}{3}{4}", bookName, bookName.EndsWith(" ") ? string.Empty : " ", number, ChapterVerseDelimiter, 0);

                            VersePointer vp = new VersePointer(verseName);
                            if (vp.IsValid)
                            {
                                result.ChapterName = string.Format("{0}{1}{2}", bookName, bookName.EndsWith(" ") ? string.Empty : " ", number);
                                bool chapterOnlyAtStartString = startIndex == -1 && nextChar != "-";    // так как пока мы не поддерживаем Откр 1-2, и так как мы найдём в таком случае только Откр 1, то не считаем это ссылкой, как ChapterOnlyAtStartString
                                if (isInBrackets && isTitle)
                                    result.ResultType = VersePointerSearchResult.SearchResultType.ExcludableChapter;
                                else
                                    result.ResultType = chapterOnlyAtStartString                                     
                                                ? VersePointerSearchResult.SearchResultType.ChapterOnlyAtStartString
                                                : VersePointerSearchResult.SearchResultType.ChapterOnly;
                                result.VersePointerEndIndex = nextBreakIndex;
                                result.VersePointerStartIndex = startIndex;
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

        private static string GetFullVerseString(string textElementValue, string verseString, ref int endIndex)
        {
            if (StringUtils.GetChar(textElementValue, endIndex) == '-')
            {
                char tempNextChar = StringUtils.GetChar(textElementValue, endIndex + 1);
                if (StringUtils.IsDigit(tempNextChar))
                {
                    verseString = string.Format("{0}-{1}",
                        verseString,
                        StringUtils.GetNextString(textElementValue, endIndex, null, out endIndex));  // чтоб учесть случай Откр 4:5-9 - чтоб определить, где заканчивается ссылка
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
        private static string ProcessVerse(Application oneNoteApp, VersePointerSearchResult searchResult, List<VersePointer> processedVerses,
            string textToChange, string textElementValue, string noteSectionGroupName, string noteSectionName, 
            string notePageName, string notePageId, string notePageContentObjectId,
            AnalyzeDepth linkDepth, VersePointerSearchResult globalChapterSearchResult, List<VersePointerSearchResult> pageChaptersSearchResult,
            bool isLink, bool isInBrackets, bool force, out int newEndVerseIndex, out HierarchySearchManager.HierarchySearchResult hierarchySearchResult)
        {
            int startVerseNameIndex = searchResult.VersePointerStartIndex;
            int endVerseNameIndex = searchResult.VersePointerEndIndex;
            bool isChapter = VersePointerSearchResult.IsChapter(searchResult.ResultType);

            newEndVerseIndex = endVerseNameIndex;
            hierarchySearchResult = HierarchySearchManager.GetHierarchyObject(
                                                oneNoteApp, SettingsManager.Instance.NotebookId_Bible, searchResult.VersePointer);
            if (hierarchySearchResult.ResultType == HierarchySearchManager.HierarchySearchResultType.Successfully)
            {
                if (hierarchySearchResult.HierarchyStage == HierarchySearchManager.HierarchyStage.ContentPlaceholder
                    || hierarchySearchResult.HierarchyStage == HierarchySearchManager.HierarchyStage.Page)  
                {
                    if (hierarchySearchResult.HierarchyStage == HierarchySearchManager.HierarchyStage.ContentPlaceholder)
                        Logger.LogMessage("Обработка стиха: {0}", searchResult.VersePointer.OriginalVerseName);
                    else
                        Logger.LogMessage("Обработка главы: {0}", textToChange);

                    if (linkDepth >= AnalyzeDepth.GetVersesLinks)
                    {
                        //if (!isLink)  // || vp.IsMultiVerse)   // например было Ин 4:17. Потом стало Ин 4:17-19. Чтоб обновлять в таком случае ссылки.
                        {
                            string link = OneNoteUtils.GenerateHref(oneNoteApp, textToChange,
                                hierarchySearchResult.HierarchyObjectInfo.PageId, hierarchySearchResult.HierarchyObjectInfo.ContentObjectId);

                           link = string.Format("<span style='font-weight:normal'>{0}</span>", link);

                            CorrectTextToChangeBoundary(textElementValue.Substring(startVerseNameIndex + 1, endVerseNameIndex - startVerseNameIndex - 1), 
                                ref startVerseNameIndex, ref endVerseNameIndex);

                            textElementValue = string.Concat(
                                textElementValue.Substring(0, startVerseNameIndex + 1),
                                link,
                                textElementValue.Substring(endVerseNameIndex));                                

                            newEndVerseIndex = startVerseNameIndex + link.Length;
                            searchResult.VersePointerEndIndex = newEndVerseIndex;
                        }

                        if (!isChapter && linkDepth >= AnalyzeDepth.Full)   // главы сразу не обрабатываем - вдруг есть стихи этих глав в текущей заметке. Вот если нет - тогда потом и обработаем
                        {
                            bool canContinue = true;
                            if (VersePointerSearchResult.IsVerse(searchResult.ResultType))
                            {
                                if (globalChapterSearchResult != null)
                                {
                                    if (globalChapterSearchResult.ResultType == VersePointerSearchResult.SearchResultType.ChapterAndVerseAtStartString)
                                    {
                                        if (!isInBrackets)
                                        {
                                            if (globalChapterSearchResult.VersePointer.IsMultiVerse)
                                                if (globalChapterSearchResult.VersePointer.IsInVerseRange(searchResult.VersePointer.Verse.GetValueOrDefault(0)))
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

                            if (canContinue)
                            {
                                if (pageChaptersSearchResult != null)
                                {
                                    if (!isInBrackets)
                                    {
                                        if (pageChaptersSearchResult.Any(pcsr =>
                                            {
                                                return pcsr.ResultType == VersePointerSearchResult.SearchResultType.ExcludableChapter
                                                        && pcsr.VersePointer.ChapterName == searchResult.VersePointer.ChapterName;
                                            }))
                                            canContinue = false;  // то есть среди исключаемых глав есть текущая
                                    }
                                }
                            }

                            if (canContinue)
                            {
                                LinkVerseToNotesPage(oneNoteApp, searchResult.VersePointer, isChapter,
                                    processedVerses, hierarchySearchResult.HierarchyObjectInfo,
                                    noteSectionGroupName, noteSectionName, notePageName, notePageId, notePageContentObjectId, force);
                            }
                        }
                    }
                }                
            }

            return textElementValue;
        }

        /// <summary>
        /// Временное решение для проблемы "неправильно выдирает текст html, когда заменяет его ссылкой".
        /// </summary>
        /// <param name="textToChangeHtml"></param>
        /// <param name="startVerseNameIndex"></param>
        /// <param name="endVerseNameIndex"></param>
        private static void CorrectTextToChangeBoundary(string textToChangeHtml, ref int startVerseNameIndex, ref int endVerseNameIndex)
        {
            if (textToChangeHtml[textToChangeHtml.Length - 1] == '>') // пока только в таком случае корректируем
            {
                int openSpanTagsCount = StringUtils.GetEntranceCount(textToChangeHtml, "<span");
                int closedSpantagsCount = StringUtils.GetEntranceCount(textToChangeHtml, "</span");

                if (closedSpantagsCount - openSpanTagsCount == 1)  // рассматриваем случай, когда на вход пришло значение, например :2</span>
                {
                    endVerseNameIndex -= "</span>".Length;
                }
            }
        }

        private static void LinkVerseToNotesPage(Application oneNoteApp, VersePointer vp, bool isChapter,
            List<VersePointer> processedVerses, HierarchySearchManager.HierarchyObjectInfo verseHierarchyObjectInfo,
            string noteSectionGroupName, string noteSectionName,
            string notePageName, string notePageId, string notePageContentObjectId, bool force)
        {
            string pageContentXml;
            XDocument versePageDocument;
            XmlNamespaceManager xnm;
            oneNoteApp.GetPageContent(verseHierarchyObjectInfo.PageId, out pageContentXml);
            versePageDocument = OneNoteUtils.GetXDocument(pageContentXml, out xnm);
            string pageName = (string)versePageDocument.Root.Attribute("name");

            string notesPageId = null;
            try
            {
                notesPageId = VerseLinkManager.FindVerseLinkPageAndCreateIfNeeded(oneNoteApp,
                    verseHierarchyObjectInfo.SectionId,
                    verseHierarchyObjectInfo.PageId, pageName, SettingsManager.Instance.PageName_Notes);
            }
            catch (Exception ex)
            {
                Logger.LogError(ex);
            }

            if (!string.IsNullOrEmpty(notesPageId))
            {
                string targetContentObjectId = UpdateNotesPage(oneNoteApp, vp, isChapter, processedVerses, verseHierarchyObjectInfo,
                        noteSectionGroupName, noteSectionName, notesPageId, notePageName, notePageId, notePageContentObjectId, force);

                string link = string.Format("<font size='2pt'>{0}</font>",
                                OneNoteUtils.GenerateHref(oneNoteApp, SettingsManager.Instance.PageName_Notes, notesPageId, targetContentObjectId));

                bool wasModified = false;

                if (isChapter)
                {
                    if (SetLinkToNotesPageForChapter(versePageDocument, link, xnm))
                        wasModified = true;                    
                }
                else
                {
                    if (SetLinkToNotesPageForVerse(versePageDocument, link, vp, verseHierarchyObjectInfo, xnm))
                        wasModified = true;                  
                }

                if (wasModified)
                    oneNoteApp.UpdatePageContent(versePageDocument.ToString());

                processedVerses.Add(vp);
            }
        }

        private static bool SetLinkToNotesPageForVerse(XDocument pageDocument, string link, VersePointer vp,
            HierarchySearchManager.HierarchyObjectInfo verseHierarchyObjectInfo, XmlNamespaceManager xnm)
        {
            bool result = false;

            //находим ячейку для заметок стиха
            XElement contentObject = pageDocument.XPathSelectElement(string.Format("//one:OE[@objectID = '{0}']",
                verseHierarchyObjectInfo.ContentObjectId), xnm);
            if (contentObject == null)
            {
                Logger.LogError("Не найдена ячейка для стиха '{0}'", vp.OriginalVerseName);
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

        private static bool SetLinkToNotesPageForChapter(XDocument pageDocument, string link, XmlNamespaceManager xnm)
        {
            bool result = false;

            XElement notesLinkElement = pageDocument.Root.XPathSelectElement("one:Outline/one:OEChildren", xnm);

            if (notesLinkElement != null && notesLinkElement.Nodes().Count() == 1)   // похоже на правду                        
            {
                notesLinkElement = notesLinkElement.XPathSelectElement(string.Format("one:OE/one:T[contains(.,'{0}')]",
                    SettingsManager.Instance.PageName_Notes), xnm);
            }
            else notesLinkElement = null;

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

        private static string UpdateNotesPage(Application oneNoteApp, VersePointer vp, bool isChapter,
            List<VersePointer> processedVerses, HierarchySearchManager.HierarchyObjectInfo verseHierarchyObjectInfo,
            string noteSectionGroupName, string noteSectionName,
            string notesPageId, string notePageName, string notePageId, string notePageContentObjectId, bool force)
        {
            string targetContentObjectId = string.Empty;
            string notesPageContentXml;
            XDocument notesPageDocument;
            XmlNamespaceManager xnm;
            XNamespace nms = XNamespace.Get(Constants.OneNoteXmlNs);
            oneNoteApp.GetPageContent(notesPageId, out notesPageContentXml);
            notesPageDocument = OneNoteUtils.GetXDocument(notesPageContentXml, out xnm);

            XElement rowElement = GetNotesRowAndCreateIfNotExists(oneNoteApp, vp, isChapter, verseHierarchyObjectInfo, notesPageDocument, xnm, nms);            

            if (rowElement != null)
            {
                AddLinkToNotePage(oneNoteApp, vp, processedVerses, rowElement, noteSectionGroupName, noteSectionName,
                    notePageName, notePageId, notePageContentObjectId, notesPageDocument, xnm, nms, force);

                targetContentObjectId = GetNotesRowObjectId(oneNoteApp, notesPageId, vp, isChapter);               
            }            

            return targetContentObjectId;
        }

        private static void AddLinkToNotePage(Application oneNoteApp, VersePointer vp, List<VersePointer> processedVerses, XElement rowElement, 
            string noteSectionGroupName, string noteSectionName,
            string notePageName, string notePageId, string notePageContentObjectId,
            XDocument notesPageDocument, XmlNamespaceManager xnm, XNamespace nms, bool force)
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
                suchNoteLink.Value = suchNoteLink.Value.Replace("\n", " ");

            string multiVerseString = GetMultiVerseString(suchNoteLink, vp, processedVerses, force);
            string verseCountString = GetVerseCountString(suchNoteLink, vp, processedVerses, force);
            string fullLinkString = string.Format("{0} <b>{1}{3}{2}</b>", link, multiVerseString, verseCountString, string.IsNullOrEmpty(multiVerseString) ? string.Empty : " ");

            if (suchNoteLink == null)  // если нет ссылки на такую же заметку
            {
                notesCellElement.Add(new XElement(nms + "OE",
                                        new XElement(nms + "List",
                                                    new XElement(nms + "Number", new XAttribute("numberSequence", 0), new XAttribute("numberFormat", "##."))),
                                        new XElement(nms + "T",
                                            new XCData(
                                                fullLinkString))));
            }
            else
            {
                suchNoteLink.Value = fullLinkString;

                if (suchNoteLink.Parent.XPathSelectElement("one:List", xnm) == null)  // почему то нет номера у строки
                    suchNoteLink.Parent.AddFirst(new XElement(nms + "List",
                                                    new XElement(nms + "Number", new XAttribute("numberSequence", 0), new XAttribute("numberFormat", "##."))));                    
            
            }

            oneNoteApp.UpdatePageContent(notesPageDocument.ToString());
        }

        private static string GetVerseCountString(XElement suchNoteLink, VersePointer vp, List<VersePointer> processedVerses, bool force)
        {
            string verseCountString = string.Empty;

            if (suchNoteLink != null)
            {
                string verseCountSearchPattern = "[";
                string verseCountEndSearchPattern = "]";

                int verseCount = 1;  // один раз уже встречается стих, так как suchNoteLink != null

                if (force && !processedVerses.Contains(vp))  // если в первый раз и force
                    verseCount = 0;

                int linkEndIndex = suchNoteLink.Value.IndexOf("</a");
                int verseCountIndex = suchNoteLink.Value.IndexOf(verseCountSearchPattern, linkEndIndex != -1 ? linkEndIndex : 0);

                if ((verseCountIndex != -1) && ((force && processedVerses.Contains(vp)) || !force))                    
                {
                    int verseCountEndIndex = suchNoteLink.Value.IndexOf(verseCountEndSearchPattern, verseCountIndex + 1);
                    if (verseCountEndIndex != -1)
                    {
                        int? prevVerseCount = StringUtils.GetStringFirstNumber(suchNoteLink.Value.Substring(verseCountIndex + verseCountSearchPattern.Length, verseCountEndIndex - verseCountIndex));
                        if (prevVerseCount.HasValue)
                        {
                            verseCount = prevVerseCount.Value + 1;
                        }
                    }
                }
                else
                {
                    verseCount++;                    
                }

                if (verseCount > 1)  // [1] не пишем
                    verseCountString = string.Format("{0}{1}{2}", verseCountSearchPattern, verseCount, verseCountEndSearchPattern);
            }

            return verseCountString;
        }

        private static string GetMultiVerseString(XElement suchNoteLink, VersePointer vp, List<VersePointer> processedVerses, bool force)
        {
            string multiVerseString = string.Empty;

            string topVerseSearchPattern = "(:";
            string topVerseEndSearchPattern = ")";
            string topVerseTextSearchPattern = string.Format(":{0}-", vp.Verse);            
            bool needSetTopVerse = false;
            int topVerseIndex = -1;
            
            if (suchNoteLink != null)
                topVerseIndex = suchNoteLink.Value.IndexOf(topVerseSearchPattern);

            if (topVerseIndex != -1)
            {
                int topVerseEndIndex = suchNoteLink.Value.IndexOf(topVerseEndSearchPattern, topVerseIndex + 1);
                if (topVerseEndIndex != -1)
                {
                    multiVerseString = suchNoteLink.Value.Substring(topVerseIndex, topVerseEndIndex - topVerseIndex + 1);

                    if (force && !processedVerses.Contains(vp))  // если в первый раз и force
                        multiVerseString = string.Empty;

                    if (vp.IsMultiVerse)
                    {
                        int topVerseTextIndex = multiVerseString.IndexOf(topVerseTextSearchPattern);

                        if (topVerseTextIndex != -1)
                        {
                            int? prevTopVerse = StringUtils.GetStringFirstNumber(
                                multiVerseString.Substring(topVerseTextIndex + topVerseTextSearchPattern.Length));
                            if (prevTopVerse.HasValue)
                            {
                                if (vp.TopVerse > prevTopVerse)
                                {
                                    needSetTopVerse = true;
                                }
                            }
                        }
                        else
                            needSetTopVerse = true;
                    }
                }
            }
            else
            {
                if (vp.IsMultiVerse)
                    needSetTopVerse = true;
            }

            if (needSetTopVerse)
                multiVerseString = string.Format("{0}{1}-{2}{3}", topVerseSearchPattern, vp.Verse, vp.TopVerse, topVerseEndSearchPattern);

            return multiVerseString;
        }        

        private static string GetNotesRowObjectId(Application oneNoteApp, string notesPageId, VersePointer vp, bool isChapter)
        {
            string result = string.Empty;

            XmlNamespaceManager xnm;
            string notesPageContentXml;
            oneNoteApp.GetPageContent(notesPageId, out notesPageContentXml);
            XDocument notesPageDocument = OneNoteUtils.GetXDocument(notesPageContentXml, out xnm);
            XElement tableElement = notesPageDocument.XPathSelectElement("//one:Outline/one:OEChildren/one:OE/one:Table", xnm);
            XElement targetElement = GetNotesRow(tableElement, vp, isChapter, xnm);

            if (targetElement != null)
                result = (string)targetElement.XPathSelectElement("one:Cell/one:OEChildren/one:OE", xnm).Attribute("objectID");

            return result;
        }

        private static XElement GetNotesRowAndCreateIfNotExists(Application oneNoteApp, VersePointer vp, bool isChapter, HierarchySearchManager.HierarchyObjectInfo verseHierarchyObjectInfo, 
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
                                                            new XElement(nms + "Column", new XAttribute("index", 1), new XAttribute("width", 500), new XAttribute("isLocked", true))
                                                                ))))));
                rootElement = notesPageDocument.XPathSelectElement("//one:Outline/one:OEChildren/one:OE", xnm);
            }

            XElement tableElement = rootElement.XPathSelectElement("one:Table", xnm);

            if (tableElement == null)
            {
                rootElement.Add(new XElement(nms + "Table", new XAttribute("bordersVisible", true)));

                tableElement = rootElement.XPathSelectElement("one:Table", xnm);
            }

            XElement rowElement = GetNotesRow(tableElement, vp, isChapter, xnm);                

            if (rowElement == null)
            {
                AddNewNotesRow(oneNoteApp, vp, isChapter, verseHierarchyObjectInfo, tableElement, xnm, nms);

                rowElement = GetNotesRow(tableElement, vp, isChapter, xnm);                
            }

            return rowElement;
        }

        private static XElement GetNotesRow(XElement tableElement, VersePointer vp, bool isChapter, XmlNamespaceManager xnm)
        {

            XElement result = !isChapter ? 
                                tableElement
                                   .XPathSelectElement(string.Format("one:Row/one:Cell[1]/one:OEChildren/one:OE/one:T[contains(.,'>:{0}<')]", vp.Verse.GetValueOrDefault(0)), xnm)
                              : tableElement
                              //      .XPathSelectElement(string.Format("one:Row/one:Cell[1]/one:OEChildren/one:OE/one:T[contains(.,'>{0}<')]", "глава"), xnm);
                                   .XPathSelectElement(string.Format("one:Row/one:Cell[1]/one:OEChildren/one:OE/one:T[normalize-space(.)='']", vp.Verse), xnm)
                                ;

            if (result != null)
                result = result.Parent.Parent.Parent.Parent;

            //if (isChapter && result != null)
            //{
            //    if (result.NodesBeforeSelf().Count() > 1)  // почему то иногда появляется в середине
            //        result = null;
            //}

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
                    XCData verseData = (XCData)row.Nodes().First();
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

        public static void DeletePageNotes(Application oneNoteApp, string sectionGroupId, string sectionId, string pageId, string pageName)
        {
            try
            {
                bool wasModified = false;
                string pageContentXml;
                XDocument notePageDocument;
                XmlNamespaceManager xnm;
                oneNoteApp.GetPageContent(pageId, out pageContentXml);
                notePageDocument = OneNoteUtils.GetXDocument(pageContentXml, out xnm);

                foreach (XElement noteTextElement in notePageDocument.Root.XPathSelectElements("//one:Table/one:Row/one:Cell[2]/one:OEChildren/one:OE/one:T", xnm))
                {
                    if (!string.IsNullOrEmpty(noteTextElement.Value))
                    {
                        if (CantainsLinkToNotesPage(noteTextElement))
                        {
                            noteTextElement.Value = string.Empty;
                            wasModified = true;
                        }
                    }
                }

                XElement chapterNotesLink = FindChapterNotesLink(notePageDocument, xnm);
                if (chapterNotesLink != null)
                {
                    oneNoteApp.DeletePageContent(pageId, (string)chapterNotesLink.Attribute("objectID"));
                    chapterNotesLink.Remove();
                    wasModified = true;
                }

                if (wasModified)  // значит есть страница заметок
                {
                    string notesPageId = null;
                    try
                    {
                        notesPageId = VerseLinkManager.FindVerseLinkPageAndCreateIfNeeded(oneNoteApp,
                            sectionId, pageId, pageName, SettingsManager.Instance.PageName_Notes);
                    }
                    catch (Exception ex)
                    {
                        Logger.LogError(ex);
                    }

                    if (!string.IsNullOrEmpty(notesPageId))
                    {
                        oneNoteApp.DeleteHierarchy(notesPageId);
                    }
                }

                if (wasModified)
                    oneNoteApp.UpdatePageContent(notePageDocument.ToString());
            }
            catch (Exception ex)
            {
                Logger.LogError("Ошибки при обработке страницы.", ex);
            }
        }

        private static XElement FindChapterNotesLink(XDocument notePageDocument, XmlNamespaceManager xnm)
        {
            foreach (XElement outline in notePageDocument.Root.XPathSelectElements("//one:Outline", xnm))
            {
                List<XElement> textElements = outline.XPathSelectElements(".//one:T", xnm).ToList();
                if (textElements.Count == 1)
                {
                    if (CantainsLinkToNotesPage(textElements.First()))
                    {
                        return outline;
                    }
                }                
            }

            return null;
        }

        private static bool CantainsLinkToNotesPage(XElement textElement)
        {
            return textElement.Value.IndexOf(string.Format(">{0}<", SettingsManager.Instance.PageName_Notes)) != -1;                
        }
    }
}
