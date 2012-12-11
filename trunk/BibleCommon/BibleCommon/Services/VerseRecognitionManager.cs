using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using BibleCommon.Helpers;
using BibleCommon.Common;

namespace BibleCommon.Services
{
    public static class VerseRecognitionManager
    {
        internal const char ChapterVerseDelimiter = ':';
        internal const int MaxVerse = 200;

        public class LinkInfo
        { 
            public enum LinkTypeEnum
            {
                None,
                LinkAfterQuickAnalyze,
                LinkAfterFullAnalyze
            }

            public LinkTypeEnum LinkType { get; set; }
            public bool ExtendedVerse { get; set; }  // вначале было Ин 1:6. Проанализировали. Потом дописали "-7". И вроде бы ссылка уже есть, но анализировать снова надо. Но при этом нужно понимать, что первый стих уже анализировали ранее.
            public bool IsLink
            {
                get
                {
                    return LinkType != LinkTypeEnum.None;
                }
            }

            public LinkInfo()
            {
                LinkType = LinkTypeEnum.None;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="textElement"></param>
        /// <param name="numberIndex"></param>
        /// <param name="number"></param>
        /// <param name="breakIndex"></param>
        /// <param name="linkInfo"></param>
        /// <param name="isInBrackets">если в скобках []</param>
        /// <param name="isExcluded">Не стоит его по дефолту анализировать</param>
        /// <returns></returns>
        public static bool CanProcessAtNumberPosition(XElement textElement, int numberIndex,
            out int number, out int textBreakIndex, out int htmlBreakIndex, out LinkInfo linkInfo, out bool isInBrackets, out bool isExcluded)
        {
            linkInfo = new LinkInfo();
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

            if (IsLikeVerse(prevChar, nextChar))
            {   
                if (int.TryParse(numberString, out number))
                {
                    if (number > 0 && number <= MaxVerse)
                    {
                        string textString;
                        if (StringUtils.IsSurroundedBy(textElement.Value, "<a", "</a", numberIndex, true, out textString))
                        {
                            if (textString.Contains(Consts.Constants.QueryParameter_QuickAnalyze))
                                linkInfo.LinkType = LinkInfo.LinkTypeEnum.LinkAfterQuickAnalyze;                            
                            else
                                linkInfo.LinkType = LinkInfo.LinkTypeEnum.LinkAfterFullAnalyze;

                            if (textString.Contains(Consts.Constants.QueryParameter_ExtendedVerse))
                                linkInfo.ExtendedVerse = true;
                        }                        

                        isInBrackets = StringUtils.IsSurroundedBy(textElement.Value, "[", "]", numberIndex, false, out textString);
                        isExcluded = StringUtils.IsSurroundedBy(textElement.Value, "{", "}", numberIndex, false, out textString);

                        return true;
                    }
                }
            }

            return false;
        }
        
        // похоже на Библейскую ссылку
        private static bool IsLikeVerse(char prevChar, char nextChar)
        {
            char[] startChars = { ChapterVerseDelimiter, '>', ',', ';', ' ', '.' };
            return (startChars.Contains(prevChar) || StringUtils.IsCharAlphabetical(prevChar))
                && !StringUtils.IsCharAlphabetical(nextChar);
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
        public static VersePointerSearchResult GetValidVersePointer(XElement textElement,
            int numberStartIndex, int numberEndIndex, int number, VersePointerSearchResult globalChapterResult, string localChapterName,
            VersePointerSearchResult prevResult, bool isLink, bool isInBrackets, bool isTitle)
        {
            VersePointerSearchResult result = new VersePointerSearchResult();            
            int nextTextBreakIndex;
            int nextHtmlBreakIndex;
            int prevTextBreakIndex;
            int prevHtmlBreakIndex;
            
            string prevChar = GetPrevStringDesirableNotSpace(textElement.Value, numberStartIndex, new string[] { ",", ";" }, 
                null, isLink, out prevTextBreakIndex, out prevHtmlBreakIndex, StringSearchIgnorance.None, StringSearchMode.SearchFirstValueChar);
            string nextChar = StringUtils.GetNextString(textElement.Value, numberEndIndex, 
                null, out nextTextBreakIndex, out nextHtmlBreakIndex, StringSearchIgnorance.None, StringSearchMode.SearchFirstValueChar);

            if (IsFullVerse(prevChar, nextChar))
            {
                result = TryToGetFullVerse(textElement, numberStartIndex, number, nextTextBreakIndex, nextHtmlBreakIndex, prevTextBreakIndex, prevHtmlBreakIndex, 
                                                prevChar, nextChar, globalChapterResult, localChapterName, prevResult, isLink, isInBrackets, isTitle);
            }

            if (result.ResultType == VersePointerSearchResult.SearchResultType.Nothing && IsSingleVerse(prevChar, nextChar))
            {
                result = TryToGetSingleVerse(textElement, numberStartIndex, number, nextTextBreakIndex, nextHtmlBreakIndex, prevTextBreakIndex, prevHtmlBreakIndex,
                                                prevChar, nextChar, globalChapterResult, localChapterName, prevResult, isLink, isInBrackets, isTitle);
            }
            
            if (result.ResultType == VersePointerSearchResult.SearchResultType.Nothing && IsChapter(prevChar, nextChar))
            {
                result = TryToGetChapter(textElement, numberStartIndex, number, nextTextBreakIndex, nextHtmlBreakIndex, prevTextBreakIndex, prevHtmlBreakIndex,
                                                prevChar, nextChar, globalChapterResult, localChapterName, prevResult, isLink, isInBrackets, isTitle);
            }

            var tryToSearchSingleVerse = false; 
            if (result.ResultType == VersePointerSearchResult.SearchResultType.Nothing && IsChapterOrVerseWithoutBookName(prevChar, nextChar))
            {                
                result = TryGetChapterOrVerseWithoutBookName(textElement, numberStartIndex, number, nextTextBreakIndex, nextHtmlBreakIndex, prevTextBreakIndex, prevHtmlBreakIndex,
                                                prevChar, nextChar, globalChapterResult, localChapterName, prevResult, isLink, isInBrackets, isTitle, out tryToSearchSingleVerse);
            }

            // конечно наверное правильнее просто поставить вызов метода TryToGetSingleVerse в конец и проверять условие 
            // result.ResultType == VersePointerSearchResult.SearchResultType.Nothing && (IsSingleVerse(prevChar, nextChar) || tryToSearchSingleVerse)
            // но почему то опасно сейчас менять очерёдность проверок. Вроде как в этом был какой-то смысл
            if (result.ResultType == VersePointerSearchResult.SearchResultType.Nothing && tryToSearchSingleVerse)
            {
                result = TryToGetSingleVerse(textElement, numberStartIndex, number, nextTextBreakIndex, nextHtmlBreakIndex, prevTextBreakIndex, prevHtmlBreakIndex,
                                                prevChar, nextChar, globalChapterResult, localChapterName, prevResult, isLink, isInBrackets, isTitle);
            }

            return result;
        }

        private static string GetPrevStringDesirableNotSpace(string s, int index, string[] searchStrings, SearchMissInfo missInfo, bool isLink, out int textBreakIndex, out int htmlBreakIndex,
            StringSearchIgnorance ignoreSpaces = StringSearchIgnorance.None, StringSearchMode searchMode = StringSearchMode.NotSpecified)
        {            
            string result = StringUtils.GetPrevString(s, index, missInfo, out textBreakIndex, out htmlBreakIndex, ignoreSpaces, searchMode);
            if (result == " ")
            {   
                int tempTextBreakIndex;
                int tempHtmlBreakIndex;
                string tempResult = StringUtils.GetPrevString(s, htmlBreakIndex, missInfo, out tempTextBreakIndex, out tempHtmlBreakIndex, ignoreSpaces, searchMode);
                if (!string.IsNullOrEmpty(tempResult) && searchStrings.Contains(tempResult))
                {
                    result = tempResult;                    
                    htmlBreakIndex = tempHtmlBreakIndex;

                    if (!isLink)  // на самом деле это не так критично (вроде бы). Но если isLink, то он не должен менять textBreakIndex, потому что обычно пробел идёт за ссылкой. В таком случае метод CorrectTextToChangeBoundary не корректно отрабатывает. А учесть пробел необходимо, чтобы удалить его в строке "Ин 1:2, 3". Но когда у нас isLink, то по идее таких пробелов и быть больше не должно
                        textBreakIndex = tempTextBreakIndex;
                }
            }

            return result;
        }

        private static string GetNextStringDesirableNotSpace(string s, int index, string[] searchStrings, SearchMissInfo missInfo, bool isLink, 
                        out int textBreakIndex, out int htmlBreakIndex, out bool spaceWasFound,
                        StringSearchIgnorance ignoreSpaces = StringSearchIgnorance.None, StringSearchMode searchMode = StringSearchMode.NotSpecified)
        {
            spaceWasFound = false;
            string result = StringUtils.GetNextString(s, index, missInfo, out textBreakIndex, out htmlBreakIndex, ignoreSpaces, searchMode);
            if (result == " ")
            {
                spaceWasFound = true;
                int tempTextBreakIndex;
                int tempHtmlBreakIndex;
                string tempResult = StringUtils.GetNextString(s, htmlBreakIndex, missInfo, out tempTextBreakIndex, out tempHtmlBreakIndex, ignoreSpaces, searchMode);
                if (!string.IsNullOrEmpty(tempResult) && searchStrings.Contains(tempResult))
                {
                    result = tempResult;
                    htmlBreakIndex = tempHtmlBreakIndex;

                    if (!isLink)  // на самом деле это не так критично (вроде бы). Но если isLink, то он не должен менять textBreakIndex, потому что обычно пробел идёт за ссылкой. В таком случае метод CorrectTextToChangeBoundary не корректно отрабатывает. А учесть пробел необходимо, чтобы удалить его в строке "Ин 1:2, 3". Но когда у нас isLink, то по идее таких пробелов и быть больше не должно
                        textBreakIndex = tempTextBreakIndex;
                }
            }

            return result;
        }

        private static VersePointerSearchResult TryGetChapterOrVerseWithoutBookName(XElement textElement, int numberStartIndex, int number, 
            int nextTextBreakIndex, int nextHtmlBreakIndex, int prevTextBreakIndex, int prevHtmlBreakIndex, string prevChar, string nextChar, 
            VersePointerSearchResult globalChapterResult, string localChapterName, 
            VersePointerSearchResult prevResult, bool isLink, bool isInBrackets, bool isTitle, out bool tryToSearchSingleVerse)
        {
            VersePointerSearchResult result = new VersePointerSearchResult();
            tryToSearchSingleVerse = false;

            int temp, temp2, endIndex = 0;
            string prevPrevChar = StringUtils.GetPrevString(textElement.Value, prevHtmlBreakIndex, null, out temp, out temp2, StringSearchIgnorance.None, StringSearchMode.SearchFirstChar);           

            if (!string.IsNullOrEmpty(localChapterName))
            {
                bool canContinue = true;                
                
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
                        canContinue = IsVersePointerFollowedByThePrevResult(textElement, prevHtmlBreakIndex, prevResult);
                    }
                }

                if (canContinue)
                {   
                    VersePointerSearchResult.SearchResultType resultType = VersePointerSearchResult.SearchResultType.Nothing;
                    string verseName = string.Empty;
                    string verseString = string.Empty;

                    if (nextChar == ChapterVerseDelimiter.ToString())
                    {
                        verseString = ExtractFullVerseString(textElement, isLink, ref nextHtmlBreakIndex, out endIndex);

                        if (!string.IsNullOrEmpty(verseString))
                        {
                            resultType = VersePointerSearchResult.SearchResultType.ChapterAndVerseWithoutBookName;
                            verseName = GetVerseName(prevResult.VersePointer.OriginalBookName, number, verseString);
                        }
                        else  // то есть вроде бы похоже на главу, а на самом деле нет стиха. Скорее всего ссылка типа "Ин 1:5,6: вот"
                        {
                            if (prevChar == "," && !prevResult.VersePointer.IsChapter)  
                                tryToSearchSingleVerse = true;                            
                        }
                    }

                    if (resultType == VersePointerSearchResult.SearchResultType.Nothing && !tryToSearchSingleVerse)
                    {
                        endIndex = nextTextBreakIndex;
                        verseString = GetFullVerseString(textElement.Value, number, isLink, ref endIndex, ref nextHtmlBreakIndex);

                        if (!string.IsNullOrEmpty(verseString))
                        {
                            resultType = isTitle && isInBrackets 
                                ? VersePointerSearchResult.SearchResultType.ExcludableChapterWithoutBookName 
                                : VersePointerSearchResult.SearchResultType.ChapterWithoutBookName;
                            verseName = GetVerseName(prevResult.VersePointer.OriginalBookName, verseString);
                        }
                    }

                    if (!string.IsNullOrEmpty(verseName) && resultType != VersePointerSearchResult.SearchResultType.Nothing)
                    {
                        VersePointer vp = new VersePointer(verseName);

                        if (vp.IsValid)
                        {
                            result.VersePointer = vp;
                            result.ResultType = resultType;
                            result.VersePointerEndIndex = endIndex;
                            result.VersePointerStartIndex = prevTextBreakIndex + 1;
                            result.VersePointerHtmlEndIndex = nextHtmlBreakIndex;
                            result.VersePointerHtmlStartIndex = prevHtmlBreakIndex;
                            result.ChapterName = GetVerseName(prevResult.VersePointer.OriginalBookName, number);
                            result.VerseString = ((resultType == VersePointerSearchResult.SearchResultType.ChapterWithoutBookName 
                                                        || resultType == VersePointerSearchResult.SearchResultType.ExcludableChapterWithoutBookName)
                                                    ? verseString
                                                    : string.Format("{0}{1}{2}", number, ChapterVerseDelimiter, verseString));
                            result.VerseStringStartsWithSpace = true;
                            result.TextElement = textElement;
                        }
                    }
                }
            }

            return result;
        }
        

        private static VersePointerSearchResult TryToGetChapter(XElement textElement, int numberStartIndex, int number,
            int nextTextBreakIndex, int nextHtmlBreakIndex, int prevTextBreakIndex, int prevHtmlBreakIndex, string prevChar, string nextChar, 
            VersePointerSearchResult globalChapterResult, string localChapterName,
            VersePointerSearchResult prevResult, bool isLink, bool isInBrackets, bool isTitle)
        {
            VersePointerSearchResult result = new VersePointerSearchResult();

            int startIndex, endIndex;

            endIndex = nextTextBreakIndex;            
            string chapterString = GetFullVerseString(textElement.Value, number, isLink, ref endIndex, ref nextHtmlBreakIndex);

            if (!string.IsNullOrEmpty(chapterString))
            {
                for (int maxMissCount = 2; maxMissCount >= 0; maxMissCount--)
                {
                    string bookName = StringUtils.GetPrevString(textElement.Value,
                        numberStartIndex, new SearchMissInfo(maxMissCount, SearchMissInfo.MissMode.CancelOnMissFound), out startIndex, out prevHtmlBreakIndex,
                        maxMissCount > 0 ? StringSearchIgnorance.IgnoreAllSpacesAndDots : StringSearchIgnorance.IgnoreFirstSpacesAndDots,
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
                            string verseName = GetVerseName(bookName, chapterString);

                            VersePointer vp = new VersePointer(verseName);
                            if (vp.IsValid)
                            {
                                result.ChapterName = GetVerseName(bookName, number);
                                bool chapterOnlyAtStartString = prevHtmlBreakIndex == -1 && nextChar != "-";    //  не считаем это ссылкой, как ChapterOnlyAtStartString
                                if (isInBrackets && isTitle)
                                    result.ResultType = VersePointerSearchResult.SearchResultType.ExcludableChapter;
                                else
                                    result.ResultType = chapterOnlyAtStartString
                                                ? VersePointerSearchResult.SearchResultType.ChapterOnlyAtStartString
                                                : VersePointerSearchResult.SearchResultType.ChapterOnly;
                                result.VersePointerEndIndex = endIndex;
                                result.VersePointerStartIndex = startIndex + 1;
                                result.VersePointerHtmlEndIndex = nextHtmlBreakIndex;
                                result.VersePointerHtmlStartIndex = prevHtmlBreakIndex;
                                result.VersePointer = vp;
                                result.TextElement = textElement;
                                result.VerseString = verseName;
                                break;
                            }
                        }
                    }
                }
            }

            return result;
        }



        private static VersePointerSearchResult TryToGetSingleVerse(XElement textElement, int numberStartIndex, int number,
            int nextTextBreakIndex, int nextHtmlBreakIndex, int prevTextBreakIndex, int prevHtmlBreakIndex, string prevChar, string nextChar,
            VersePointerSearchResult globalChapterResult, string localChapterName,
            VersePointerSearchResult prevResult, bool isLink, bool isInBrackets, bool isTitle)
        {
            VersePointerSearchResult result = new VersePointerSearchResult();

            int temp, temp2, endIndex;
            string prevPrevChar = StringUtils.GetPrevString(textElement.Value, prevHtmlBreakIndex, null, out temp, out temp2, StringSearchIgnorance.None, StringSearchMode.SearchFirstChar);
            string globalChapterName = globalChapterResult != null ? globalChapterResult.ChapterName : string.Empty;
            string chapterName = !string.IsNullOrEmpty(globalChapterName) ? globalChapterName : localChapterName;

            if (!string.IsNullOrEmpty(chapterName))
            {
                bool canContinue = true;

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
                            canContinue = IsVersePointerFollowedByThePrevResult(textElement, prevHtmlBreakIndex, prevResult);
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
                        if (globalChapterResult != null && globalChapterResult.DepthLevel > OneNoteUtils.GetDepthLevel(textElement))  // то есть глава находится глубже (правее), чем найденный стих. должно быть либо на одном уровне, либо левее.
                            canContinue = false;
                    }
                }

                if (canContinue)
                {
                    if (!string.IsNullOrEmpty(chapterName))
                    {                        
                        endIndex = nextTextBreakIndex;
                        var verseString = GetFullVerseString(textElement.Value, number, isLink, ref endIndex, ref nextHtmlBreakIndex);

                        if (!string.IsNullOrEmpty(verseString))
                        {
                            string verseName = string.Format("{0}{1}{2}", chapterName, ChapterVerseDelimiter, verseString);

                            VersePointer vp = new VersePointer(verseName);
                            if (vp.IsValid)
                            {
                                result.VersePointer = vp;
                                result.ResultType = resultType;
                                result.VersePointerEndIndex = endIndex;
                                result.VersePointerStartIndex = prevChar == ChapterVerseDelimiter.ToString() ? prevTextBreakIndex : prevTextBreakIndex + 1;
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

            return result;
        }


        private static VersePointerSearchResult TryToGetFullVerse(XElement textElement, int numberStartIndex, int number,
            int nextTextBreakIndex, int nextHtmlBreakIndex, int prevTextBreakIndex, int prevHtmlBreakIndex, string prevChar, string nextChar,
            VersePointerSearchResult globalChapterResult, string localChapterName,
            VersePointerSearchResult prevResult, bool isLink, bool isInBrackets, bool isTitle)
        {
            VersePointerSearchResult result = new VersePointerSearchResult();

            int startIndex, endIndex;

            string verseString = ExtractFullVerseString(textElement, isLink, ref nextHtmlBreakIndex, out endIndex);

            if (!string.IsNullOrEmpty(verseString))
            {
                for (int maxMissCount = 2; maxMissCount >= 0; maxMissCount--)
                {
                    string bookName = StringUtils.GetPrevString(textElement.Value,
                        numberStartIndex, new SearchMissInfo(maxMissCount, SearchMissInfo.MissMode.CancelOnMissFound), out startIndex, out prevHtmlBreakIndex,
                        maxMissCount > 0 ? StringSearchIgnorance.IgnoreAllSpacesAndDots : StringSearchIgnorance.IgnoreFirstSpacesAndDots,
                        StringSearchMode.SearchText);

                    if (!string.IsNullOrEmpty(bookName) && !string.IsNullOrEmpty(bookName.Trim()))
                    {
                        char prevPrevChar = StringUtils.GetChar(textElement.Value, prevHtmlBreakIndex);
                        if (!(StringUtils.IsCharAlphabetical(prevPrevChar) || StringUtils.IsDigit(prevPrevChar)))
                        {
                            bookName = bookName.Trim();
                            string verseName = GetVerseName(bookName, number, verseString);

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

            return result;
        }

        private static bool IsFullVerse(string prevChar, string nextChar)
        {
            return nextChar == ChapterVerseDelimiter.ToString();
        }

        private static bool IsSingleVerse(string prevChar, string nextChar)
        {
            return (prevChar == ChapterVerseDelimiter.ToString() || prevChar == ",") && (nextChar != ChapterVerseDelimiter.ToString());
        }

        private static bool IsChapterOrVerseWithoutBookName(string prevChar, string nextChar)
        {
            string[] startChars = { ";", "," };
            return startChars.Contains(prevChar);                
        }

        private static bool IsChapter(string prevChar, string nextChar)
        {
            string[] endChars = { ")", "]", "}", ",", ".", "?", "!", ";", "-", "&", ":" };   // двоеточие добавлено потому, что могут быть ссылки типа "Ин 1: вот". Всё равно, если это была нормальная ссылка, он до сюда не дойдёт
            return string.IsNullOrEmpty(nextChar.Trim()) || endChars.Contains(nextChar);                
        }

        private static bool IsVersePointerFollowedByThePrevResult(XElement textElement, int prevHtmlBreakIndex, VersePointerSearchResult prevResult)
        {
            bool result = true;

            if (prevHtmlBreakIndex < prevResult.VersePointerHtmlEndIndex)  // если предыдущий символ ещё относится к предыдущему результату
                result = false;
            else
            {
                string stringBetweenThisAndLastResultHTML = textElement.Value
                                                .Substring(prevResult.VersePointerHtmlEndIndex, prevHtmlBreakIndex - prevResult.VersePointerHtmlEndIndex + 1);
                if (StringUtils.GetText(stringBetweenThisAndLastResultHTML).Length > 1)  // то есть не рядом был предыдущий результат
                    result = false;
            }

            return result;
        }

        private static string GetVerseName(string bookName, object chapter)
        {
            return string.Format("{0}{1}{2}", bookName, bookName.EndsWith(" ") ? string.Empty : " ", chapter);
        }

        private static string GetVerseName(string bookName, int chapter, object verseString)
        {
            return string.Format("{0}{1}{2}", GetVerseName(bookName, chapter), ChapterVerseDelimiter, verseString);
        }

        /// <summary>
        /// Для случаев, когда мы нашли номер главы, и нам надо получить номер стиха
        /// </summary>
        /// <param name="textElement"></param>
        /// <param name="nextHtmlBreakIndex"></param>
        /// <param name="endIndex"></param>
        /// <returns></returns>
        private static string ExtractFullVerseString(XElement textElement, bool isLink, ref int nextHtmlBreakIndex, out int endIndex)
        {
            string verseString = StringUtils.GetNextString(textElement.Value, nextHtmlBreakIndex, null, out endIndex, out nextHtmlBreakIndex);
            int verseNumber;
            if (int.TryParse(verseString, out verseNumber))
            {
                if (verseNumber <= MaxVerse && verseNumber > 0)
                {
                    verseString = GetFullVerseString(textElement.Value, verseNumber, isLink, ref endIndex, ref nextHtmlBreakIndex);
                    return verseString;
                }
            }

            return string.Empty;
        }

        /// <summary>
        /// Для случаев, когда у нас уже есть номер стиха, а нам надо получить полный номер стиха (например, "6-7")
        /// </summary>
        /// <param name="textElementValue"></param>
        /// <param name="verseString"></param>
        /// <param name="endIndex"></param>
        /// <param name="nextHtmlBreakIndex"></param>
        /// <returns></returns>
        private static string GetFullVerseString(string textElementValue, int verseNumber, bool isLink, ref int endIndex, ref int nextHtmlBreakIndex)
        {
            var verseString = verseNumber.ToString();

            int tempEndIndex, tempNextHtmlBreakIndex, tempTopVerse;
            bool spaceWasFound;

            string anotherKindOfDash = char.ConvertFromUtf32(8209);

            string firstNextChar = GetNextStringDesirableNotSpace(textElementValue, nextHtmlBreakIndex - 1, new string[] { "-", anotherKindOfDash },
                null, isLink, out tempEndIndex, out tempNextHtmlBreakIndex, out spaceWasFound, StringSearchIgnorance.None, StringSearchMode.SearchFirstValueChar);

            if (firstNextChar == "-" || firstNextChar == anotherKindOfDash)   
            {
                string nextChar = StringUtils.GetNextString(textElementValue, tempNextHtmlBreakIndex, null, out tempEndIndex, out tempNextHtmlBreakIndex, StringSearchIgnorance.IgnoreFirstSpaces);

                if (int.TryParse(nextChar, out tempTopVerse))
                {
                    if (!(tempTopVerse > MaxVerse))
                    {
                        var afterChar = StringUtils.GetChar(textElementValue, tempNextHtmlBreakIndex);
                        if (tempTopVerse > verseNumber
                            || afterChar == ChapterVerseDelimiter) // если строка типа Ин 1:2-3:4
                        {
                            if (!(afterChar != ChapterVerseDelimiter && spaceWasFound))
                            {
                                verseString = string.Format("{0}-{1}", verseString, nextChar.Trim());
                                endIndex = tempEndIndex;
                                nextHtmlBreakIndex = tempNextHtmlBreakIndex;

                                if (afterChar == ChapterVerseDelimiter)
                                {
                                    nextChar = StringUtils.GetNextString(textElementValue, nextHtmlBreakIndex, null, out tempEndIndex, out tempNextHtmlBreakIndex);

                                    if (int.TryParse(nextChar, out tempTopVerse))
                                    {
                                        verseString += ChapterVerseDelimiter + nextChar;
                                        endIndex = tempEndIndex;
                                        nextHtmlBreakIndex = tempNextHtmlBreakIndex;
                                    }
                                }
                            }
                        }                        
                    }
                }
                else
                {
                    if (!string.IsNullOrEmpty(nextChar) && !spaceWasFound && !nextChar.StartsWith(" ")) // чтобы отсечь варианты типа 1 Кор 2:3; 2-е Кор 3:4
                        verseString = string.Empty;
                }
            }

            return verseString;
        }
    }
}
