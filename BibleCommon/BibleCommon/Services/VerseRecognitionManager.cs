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
        private static char[] _chapterVerseDelimiter;
        private static char[] _startVerseChars;
        private static readonly object _locker = new object();

        internal const int MaxVerse = 200;
        public const char DefaultChapterVerseDelimiter = ':';

        public static char[] GetChapterVerseDelimiters()
        {            
            if (_chapterVerseDelimiter == null)
            {
                lock (_locker)
                {
                    if (_chapterVerseDelimiter == null)
                    {
                        var chars = new List<char>() { ':' };
                        if (SettingsManager.Instance.UseCommaDelimiter)
                            chars.Add(',');

                        _chapterVerseDelimiter = chars.ToArray();
                    }
                }
            }

            return _chapterVerseDelimiter;
        }


        public static char[] GetStartVerseChars()
        {
            if (_startVerseChars == null)
            {
                lock (_locker)
                {
                    if (_startVerseChars == null)
                    {
                        _startVerseChars = new List<char>(GetChapterVerseDelimiters()) { ',', '>', ';', ' ', '.' }.Distinct().ToArray();
                    }
                }
            }

            return _startVerseChars;
        }

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

        public struct VerseScopeInfo
        {
            public bool IsInBrackets { get; set; }
            public bool IsExcluded { get; set; }
            public bool IsImportantVerse { get; set; }
        }


        private static HashSet<string> _endCharsOfChapterLink;
        static VerseRecognitionManager()
        {
            _endCharsOfChapterLink = new HashSet<string>() { "(", ")", "[", "]", "{", "}", ",", ".", "?", "!", ";", "&", ":", "*" };   // двоеточие добавлено потому, что могут быть ссылки типа "Ин 1: вот". Всё равно, если это была нормальная ссылка, он до сюда не дойдёт
                                                                                                                                       // открывающие скобки добавлены, так как может быть что-то похожее на "Отк 5(синодальный перевод)"
            foreach (var dash in VerseNumber.Dashes)
                _endCharsOfChapterLink.Add(dash.ToString());
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
            out int number, out int textBreakIndex, out int htmlBreakIndex, out LinkInfo linkInfo, out VerseScopeInfo verseScopeInfo)
        {
            linkInfo = new LinkInfo();
            verseScopeInfo = new VerseScopeInfo();            
            number = -1;
            textBreakIndex = -1;
            htmlBreakIndex = -1;           

            if (numberIndex == 0)  // не может начинаться с цифры
                return false;

            char prevChar = StringUtils.GetChar(textElement.Value, numberIndex - 1);
            var numberStringResult = StringSearcher.SearchInString(textElement.Value, numberIndex - 1, StringSearcher.SearchDirection.Forward);
            textBreakIndex = numberStringResult.TextBreakIndex;
            htmlBreakIndex = numberStringResult.HtmlBreakIndex;

            var nextChar = StringUtils.GetChar(textElement.Value, htmlBreakIndex);

            if (IsLikeVerse(prevChar, nextChar))
            {
                if (int.TryParse(numberStringResult.FoundString, out number))
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

                        verseScopeInfo.IsInBrackets = StringUtils.IsSurroundedBy(textElement.Value, "[", "]", numberIndex, false, out textString);
                        verseScopeInfo.IsExcluded = StringUtils.IsSurroundedBy(textElement.Value, "{", "}", numberIndex, false, out textString);
                        verseScopeInfo.IsImportantVerse = StringUtils.IsSurroundedBy(textElement.Value, "*", "*", numberIndex, false, out textString);

                        return true;
                    }
                }
            }

            return false;
        }
        
        // похоже на Библейскую ссылку
        private static bool IsLikeVerse(char prevChar, char nextChar)
        {
            return (GetStartVerseChars().Contains(prevChar) || StringUtils.IsCharAlphabetical(prevChar))
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
            VersePointerSearchResult prevResult, bool isLink, VerseScopeInfo verseScopeInfo, bool isTitle)
        {
            VersePointerSearchResult result = new VersePointerSearchResult();            
            bool spaceWasFound;
            
            var prevCharResult = GetPrevOrNextStringDesirableNotSpace(textElement.Value, numberStartIndex, new char[] { ',', ';' },
                isLink, StringSearcher.SearchDirection.Back, StringSearcher.StringSearchMode.SearchFirstDelimiterChar, null, null, out spaceWasFound);
            var nextCharResult = StringSearcher.SearchInString(textElement.Value, numberEndIndex, StringSearcher.SearchDirection.Forward, StringSearcher.StringSearchMode.SearchFirstDelimiterChar);

            // ужасно некрасивый код! шесть раз передаём одинаково аргументы! Почему вообще это статический класс??????????
            var prevChar = prevCharResult.FoundString;
            var prevTextBreakIndex = prevCharResult.TextBreakIndex;
            var prevHtmlBreakIndex = prevCharResult.HtmlBreakIndex;

            var nextChar = nextCharResult.FoundString;
            var nextTextBreakIndex = nextCharResult.TextBreakIndex;
            var nextHtmlBreakIndex = nextCharResult.HtmlBreakIndex;

            if (IsFullVerse(prevCharResult.FoundString, nextCharResult.FoundString))
            {
                result = TryToGetFullVerse(textElement, numberStartIndex, number, nextTextBreakIndex, nextHtmlBreakIndex, prevCharResult.TextBreakIndex, prevCharResult.HtmlBreakIndex,
                                                prevChar, nextChar, globalChapterResult, localChapterName, prevResult, isLink, verseScopeInfo, isTitle);                
            }

            if (result.ResultType == VersePointerSearchResult.SearchResultType.Nothing && IsSingleVerse(prevChar, nextChar) && prevCharResult.WasFoundDigits.GetValueOrDefault(true))  // в последней проверке убеждаемся, что в найденной строке были только цифры (и разделители), но не буквы
            {
                result = TryToGetSingleVerse(textElement, numberStartIndex, number, nextTextBreakIndex, nextHtmlBreakIndex, prevCharResult.TextBreakIndex, prevCharResult.HtmlBreakIndex,
                                                prevChar, nextChar, globalChapterResult, localChapterName, prevResult, isLink, verseScopeInfo, isTitle);
            }

            if (result.ResultType == VersePointerSearchResult.SearchResultType.Nothing && IsChapter(prevChar, nextChar))
            {
                result = TryToGetChapter(textElement, numberStartIndex, number, nextTextBreakIndex, nextHtmlBreakIndex, prevCharResult.TextBreakIndex, prevCharResult.HtmlBreakIndex,
                                                prevChar, nextChar, globalChapterResult, localChapterName, prevResult, isLink, verseScopeInfo, isTitle);
            }

            var tryToSearchSingleVerse = false;
            if (result.ResultType == VersePointerSearchResult.SearchResultType.Nothing && IsChapterOrVerseWithoutBookName(prevChar, nextChar) && prevCharResult.WasFoundDigits.GetValueOrDefault(true))
            {
                result = TryGetChapterOrVerseWithoutBookName(textElement, numberStartIndex, number, nextTextBreakIndex, nextHtmlBreakIndex, prevCharResult.TextBreakIndex, prevCharResult.HtmlBreakIndex,
                                                prevChar, nextChar, globalChapterResult, localChapterName, prevResult, isLink, verseScopeInfo, isTitle, out tryToSearchSingleVerse);
            }

            // конечно наверное правильнее просто поставить вызов метода TryToGetSingleVerse в конец и проверять условие 
            // result.ResultType == VersePointerSearchResult.SearchResultType.Nothing && (IsSingleVerse(prevChar, nextChar) || tryToSearchSingleVerse)
            // но почему то опасно сейчас менять очерёдность проверок. Вроде как в этом был какой-то смысл
            if (result.ResultType == VersePointerSearchResult.SearchResultType.Nothing && tryToSearchSingleVerse)
            {
                result = TryToGetSingleVerse(textElement, numberStartIndex, number, nextTextBreakIndex, nextHtmlBreakIndex, prevTextBreakIndex, prevHtmlBreakIndex,
                                                prevChar, nextChar, globalChapterResult, localChapterName, prevResult, isLink, verseScopeInfo, isTitle);
            }

            return result;
        }

        private static StringSearcher.StringSearchResult GetPrevOrNextStringDesirableNotSpace(string s, int index, char[] searchChars, bool isLink,
            StringSearcher.SearchDirection direction, StringSearcher.StringSearchMode searchMode, StringSearcher.SearchMissInfo missInfo, StringSearcher.SearchIgnoringInfo ignoringInfo,  
            out bool spaceWasFound)
        {
            spaceWasFound = false;
            var result = StringSearcher.SearchInString(s, index, direction, searchMode, missInfo, ignoringInfo);
            if (result.FoundString == " ")
            {
                spaceWasFound = true;
                var tempResult = StringSearcher.SearchInString(s, result.HtmlBreakIndex, direction, searchMode, missInfo, ignoringInfo);
                if (!string.IsNullOrEmpty(tempResult.FoundString) && searchChars.Contains(tempResult.FoundString.FirstOrDefault()))
                {
                    result.FoundString = tempResult.FoundString;
                    result.HtmlBreakIndex = tempResult.HtmlBreakIndex;

                    if (!isLink)  // на самом деле это не так критично (вроде бы). Но если isLink, то он не должен менять textBreakIndex, потому что обычно пробел идёт за ссылкой. В таком случае метод CorrectTextToChangeBoundary не корректно отрабатывает. А учесть пробел необходимо, чтобы удалить его в строке "Ин 1:2, 3". Но когда у нас isLink, то по идее таких пробелов и быть больше не должно
                        result.TextBreakIndex = tempResult.TextBreakIndex;
                }
            }

            return result;
        }

        private static VersePointerSearchResult TryGetChapterOrVerseWithoutBookName(XElement textElement, int numberStartIndex, int number, 
            int nextTextBreakIndex, int nextHtmlBreakIndex, int prevTextBreakIndex, int prevHtmlBreakIndex, string prevChar, string nextChar, 
            VersePointerSearchResult globalChapterResult, string localChapterName,
            VersePointerSearchResult prevResult, bool isLink, VerseScopeInfo verseScopeInfo, bool isTitle, out bool tryToSearchSingleVerse)
        {
            VersePointerSearchResult result = new VersePointerSearchResult();
            tryToSearchSingleVerse = false;

            int temp, endIndex = 0;
            string prevPrevChar = StringSearcher.SearchInString(textElement.Value, prevHtmlBreakIndex, StringSearcher.SearchDirection.Back, StringSearcher.StringSearchMode.SearchFirstChar).FoundString;           

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
                    var resultType = VersePointerSearchResult.SearchResultType.Nothing;
                    var verseName = string.Empty;
                    var verseString = string.Empty;

                    if (GetChapterVerseDelimiters().Contains(StringUtils.GetChar(nextChar, 0)))
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
                            resultType = isTitle && verseScopeInfo.IsInBrackets 
                                ? VersePointerSearchResult.SearchResultType.ExcludableChapterWithoutBookName 
                                : VersePointerSearchResult.SearchResultType.ChapterWithoutBookName;
                            verseName = GetVerseName(prevResult.VersePointer.OriginalBookName, verseString);
                        }
                    }

                    if (!string.IsNullOrEmpty(verseName) && resultType != VersePointerSearchResult.SearchResultType.Nothing)
                    {
                        VersePointer vp = new VersePointer(verseName);

                        if (vp.IsValid && vp.IsExisting.GetValueOrDefault(true))
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
                                                    : string.Format("{0}{1}{2}", number, DefaultChapterVerseDelimiter, verseString));
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
            VersePointerSearchResult prevResult, bool isLink, VerseScopeInfo verseScopeInfo, bool isTitle)
        {
            VersePointerSearchResult result = new VersePointerSearchResult();

            int startIndex, endIndex;

            endIndex = nextTextBreakIndex;            
            string chapterString = GetFullVerseString(textElement.Value, number, isLink, ref endIndex, ref nextHtmlBreakIndex);

            if (!string.IsNullOrEmpty(chapterString))
            {
                StringSearcher.StringSearchResult bookNameResult = null;

                for (int counter = 2; counter >= -2; counter--)
                {
                    if (counter == 1 && !bookNameResult.MissFound)  // то есть если в предыдущий раз мы не встретили никакого miss символа (а мы готовы были аж два miss символа взять), то зачем опять искать (уже будучи готовыми взять только один miss символ)?!
                        continue;

                    bookNameResult = StringSearcher.SearchInString(
                                            textElement.Value,
                                            numberStartIndex,
                                            StringSearcher.SearchDirection.Back,
                                            StringSearcher.StringSearchMode.SearchText,
                                            new StringSearcher.SearchMissInfo(counter > 0 
                                                                    ? counter 
                                                                    : 0, 
                                                               StringSearcher.SearchMissInfo.MissMode.CancelOnMissFound),                                             
                                            GetStringSearchIgnoringInfo(counter));                    

                    var bookName = bookNameResult.FoundString;
                    startIndex = bookNameResult.TextBreakIndex;
                    prevHtmlBreakIndex = bookNameResult.HtmlBreakIndex;                    

                    if (!string.IsNullOrEmpty(bookName))
                    {
                        bookName = bookName.Trim(' ');
                        if (!string.IsNullOrEmpty(bookName))
                        {
                            bookName = TrimBookName(bookName, verseScopeInfo, textElement, ref startIndex);

                            char prevPrevChar = StringUtils.GetChar(textElement.Value, prevHtmlBreakIndex);
                            if (!(StringUtils.IsCharAlphabetical(prevPrevChar) || StringUtils.IsDigit(prevPrevChar)))
                            {
                                bookName = bookName.Trim(' ');
                                string verseName = GetVerseName(bookName, chapterString);

                                VersePointer vp = new VersePointer(verseName);
                                if (vp.IsValid && vp.IsExisting.GetValueOrDefault(true))
                                {
                                    result.ChapterName = GetVerseName(bookName, number);
                                    bool chapterOnlyAtStartString = prevHtmlBreakIndex == -1 && !VerseNumber.Dashes.Contains(nextChar.FirstOrDefault());    //  не считаем это ссылкой, как ChapterOnlyAtStartString
                                    if (verseScopeInfo.IsInBrackets && isTitle)
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
            }

            return result;
        }        

        private static VersePointerSearchResult TryToGetSingleVerse(XElement textElement, int numberStartIndex, int number,
            int nextTextBreakIndex, int nextHtmlBreakIndex, int prevTextBreakIndex, int prevHtmlBreakIndex, string prevChar, string nextChar,
            VersePointerSearchResult globalChapterResult, string localChapterName,
            VersePointerSearchResult prevResult, bool isLink, VerseScopeInfo verseScopeInfo, bool isTitle)
        {
            VersePointerSearchResult result = new VersePointerSearchResult();

            int temp, endIndex;
            var prevPrevChar = StringSearcher.SearchInString(textElement.Value, prevHtmlBreakIndex, StringSearcher.SearchDirection.Back, StringSearcher.StringSearchMode.SearchFirstChar).FoundString;
            string globalChapterName = globalChapterResult != null ? globalChapterResult.ChapterName : string.Empty;
            string chapterName = !string.IsNullOrEmpty(globalChapterName) ? globalChapterName : localChapterName;

            if (!string.IsNullOrEmpty(chapterName))
            {
                bool canContinue = true;

                var prevCharIsDefaultChapterVerseDelimiter = prevChar == DefaultChapterVerseDelimiter.ToString();

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
                                   || VersePointerSearchResult.IsVerseWithoutChapter(prevResult.ResultType)))  // если только что до этого мы нашли стих, тогда запятая - это разделитель стихов, иначе здесь может иметься ввиду разделение глав, которое пока не поддерживаем
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
                else if (prevCharIsDefaultChapterVerseDelimiter)    // надо проверить, чтоб предыдущий символ не был цифрой или буквой
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
                            string verseName = string.Format("{0}{1}{2}", chapterName, DefaultChapterVerseDelimiter, verseString);

                            VersePointer vp = new VersePointer(verseName);
                            if (vp.IsValid && vp.IsExisting.GetValueOrDefault(true))
                            {
                                result.VersePointer = vp;
                                result.ResultType = resultType;
                                result.VersePointerEndIndex = endIndex;
                                result.VersePointerStartIndex = prevCharIsDefaultChapterVerseDelimiter ? prevTextBreakIndex : prevTextBreakIndex + 1;
                                result.VersePointerHtmlEndIndex = nextHtmlBreakIndex;
                                result.VersePointerHtmlStartIndex = prevHtmlBreakIndex;
                                result.VerseString = (prevCharIsDefaultChapterVerseDelimiter ? prevChar : string.Empty) + verseString;
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
            VersePointerSearchResult prevResult, bool isLink, VerseScopeInfo verseScopeInfo, bool isTitle)
        {
            VersePointerSearchResult result = new VersePointerSearchResult();

            int startIndex, endIndex;

            string verseString = ExtractFullVerseString(textElement, isLink, ref nextHtmlBreakIndex, out endIndex);

            if (!string.IsNullOrEmpty(verseString))
            {
                StringSearcher.StringSearchResult bookNameResult = null;

                for (int counter = 2; counter >= -2; counter--)
                {
                    if (counter == 1 && !bookNameResult.MissFound)  // то есть если в предыдущий раз мы не встретили никакого miss символа (а мы готовы были аж два miss символа взять), то зачем опять искать (уже будучи готовыми взять только один miss символ)?!
                        continue;

                    bookNameResult = StringSearcher.SearchInString(
                                            textElement.Value,
                                            numberStartIndex,
                                            StringSearcher.SearchDirection.Back,
                                            StringSearcher.StringSearchMode.SearchText,
                                            new StringSearcher.SearchMissInfo(counter > 0
                                                                    ? counter
                                                                    : 0,
                                                               StringSearcher.SearchMissInfo.MissMode.CancelOnMissFound),
                                            GetStringSearchIgnoringInfo(counter));

                    var bookName = bookNameResult.FoundString;
                    startIndex = bookNameResult.TextBreakIndex;
                    prevHtmlBreakIndex = bookNameResult.HtmlBreakIndex;                  

                    if (!string.IsNullOrEmpty(bookName))
                    {
                        bookName = bookName.Trim(' ');
                        if (!string.IsNullOrEmpty(bookName))
                        {
                            bookName = TrimBookName(bookName, verseScopeInfo, textElement, ref startIndex);

                            char prevPrevChar = StringUtils.GetChar(textElement.Value, prevHtmlBreakIndex);
                            if (!(StringUtils.IsCharAlphabetical(prevPrevChar) || StringUtils.IsDigit(prevPrevChar)))
                            {
                                bookName = bookName.Trim(' ');
                                string verseName = GetVerseName(bookName, number, verseString);

                                var vp = new VersePointer(verseName);
                                if (vp.IsValid && vp.IsExisting.GetValueOrDefault(true))
                                {
                                    if (!(
                                        SettingsManager.Instance.UseCommaDelimiter
                                        && vp.Chapter > 1
                                        && SettingsManager.Instance.CanUseBibleContent
                                        && SettingsManager.Instance.CurrentBibleContentCached.BookHasOnlyOneChapter(vp.ToSimpleVersePointer())))    // например, "Иуд 14,15"
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

            return result;
        }

        private static StringSearcher.SearchIgnoringInfo GetStringSearchIgnoringInfo(int counter)
        {
            return counter > 0
                        ? new StringSearcher.SearchIgnoringInfo(int.MaxValue, StringSearcher.SearchIgnoringInfo.IgnoringMode.IgnoreSpacesAndDots)
                        : new StringSearcher.SearchIgnoringInfo(3 + counter, StringSearcher.SearchIgnoringInfo.IgnoringMode.IgnoreSpacesAndDots);                       
        }

        private static bool IsFullVerse(string prevChar, string nextChar)
        {
            return GetChapterVerseDelimiters().Contains(StringUtils.GetChar(nextChar, 0));
        }

        private static bool IsSingleVerse(string prevChar, string nextChar)
        {            
            return (prevChar == DefaultChapterVerseDelimiter.ToString() || prevChar == ",") && (nextChar != DefaultChapterVerseDelimiter.ToString());
        }

        private static bool IsChapterOrVerseWithoutBookName(string prevChar, string nextChar)
        {
            string[] startChars = { ";", "," };
            return startChars.Contains(prevChar);                
        }

        private static bool IsChapter(string prevChar, string nextChar)
        {      
            return string.IsNullOrEmpty(nextChar.Trim()) || _endCharsOfChapterLink.Contains(nextChar);                
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
            return string.Format("{0}{1}{2}", GetVerseName(bookName, chapter), DefaultChapterVerseDelimiter, verseString);
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
            int verseNumber;

            var verseStringResult = StringSearcher.SearchInString(textElement.Value, nextHtmlBreakIndex, StringSearcher.SearchDirection.Forward); 
            var verseString = verseStringResult.FoundString;
            endIndex = verseStringResult.TextBreakIndex;
            nextHtmlBreakIndex = verseStringResult.HtmlBreakIndex;

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

            int localTopVerse;
            bool spaceWasFound;

            var firstNextCharResult = GetPrevOrNextStringDesirableNotSpace(textElementValue, nextHtmlBreakIndex - 1, VerseNumber.Dashes,
                isLink, StringSearcher.SearchDirection.Forward, StringSearcher.StringSearchMode.SearchFirstDelimiterChar, null, null, out spaceWasFound);             

            if (VerseNumber.Dashes.Contains(firstNextCharResult.FoundString.FirstOrDefault()))   
            {
                var nextStringResult = StringSearcher.SearchInString(textElementValue, firstNextCharResult.HtmlBreakIndex, StringSearcher.SearchDirection.Forward, StringSearcher.StringSearchMode.NotSpecified, null,
                                                                 new StringSearcher.SearchIgnoringInfo(1, StringSearcher.SearchIgnoringInfo.IgnoringMode.IgnoreSpaces));              

                if (int.TryParse(nextStringResult.FoundString, out localTopVerse))
                {
                    if (!(localTopVerse > MaxVerse))
                    {
                        var afterChar = StringUtils.GetChar(textElementValue, nextStringResult.HtmlBreakIndex);
                        if (localTopVerse > verseNumber
                           || afterChar == DefaultChapterVerseDelimiter) // если строка типа Ин 1:2-3:4
                        {
                            if (!(afterChar != DefaultChapterVerseDelimiter && spaceWasFound))
                            {
                                verseString = string.Format("{0}-{1}", verseString, nextStringResult.FoundString.Trim());
                                endIndex = nextStringResult.TextBreakIndex;
                                nextHtmlBreakIndex = nextStringResult.HtmlBreakIndex;

                                if (afterChar == DefaultChapterVerseDelimiter)
                                {
                                    nextStringResult = StringSearcher.SearchInString(textElementValue, nextHtmlBreakIndex, StringSearcher.SearchDirection.Forward);

                                    if (int.TryParse(nextStringResult.FoundString, out localTopVerse))
                                    {
                                        verseString += DefaultChapterVerseDelimiter + nextStringResult.FoundString;
                                        endIndex = nextStringResult.TextBreakIndex;
                                        nextHtmlBreakIndex = nextStringResult.HtmlBreakIndex;
                                    }
                                }
                            }
                        }                        
                    }
                }
                else
                {
                    if (!string.IsNullOrEmpty(nextStringResult.FoundString) && !spaceWasFound && !nextStringResult.FoundString.StartsWith(" ") && nextStringResult.FoundString.Length == 1) // чтобы отсечь варианты типа 1 Кор 2:3; 2-е Кор 3:4
                        verseString = string.Empty;
                }
            }

            return verseString;
        }


        private static string TrimBookName(string bookName, VerseScopeInfo verseScopeInfo, XElement textElement, ref int startIndex)
        {
            if (verseScopeInfo.IsInBrackets)
            {
                if (textElement.Value[startIndex + 1] == '[')  // временное решение проблемы: когда указана в заголовке только глава в квадратных скобках, то первая скобка удаляется 
                    startIndex++;

                bookName = bookName.Trim('[', ']');
            }

            if (verseScopeInfo.IsExcluded)
            {
                if (textElement.Value[startIndex + 1] == '{')  
                    startIndex++;

                bookName = bookName.Trim('{', '}');
            }

            if (verseScopeInfo.IsImportantVerse)
            {
                if (textElement.Value[startIndex + 1] == '*')
                    startIndex++;

                bookName = bookName.Trim('*');
            }

            return bookName;
        }
    }
}
