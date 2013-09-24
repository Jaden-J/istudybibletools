using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BibleCommon.Helpers
{
    public class StringSearcher
    {
        #region Helper Classes

        public enum SearchDirection
        {
            Forward,
            Back
        }

        public enum StringSearchMode
        {
            NotSpecified,
            SearchNumber,
            SearchText,
            SearchFirstValueChar,   // получить первый значащий символ-разделитель (не цифра, и не симфол алфавита, не равный '<' и '>')
            SearchFirstChar         // получить первый любой символ (не равный '<' и '>')
        }

        public class SearchMissInfo
        {
            public enum MissMode
            {
                CancelOnNextMiss,        // выходим, когда количество промахов превысило лемит
                CancelOnMissFound      // как только нашли неподходящий символ - добавляем его к результативной строке и выходим
            }

            public int? MaxMissCount { get; set; }
            public MissMode Mode { get; set; }

            public SearchMissInfo(int? maxMissCount, MissMode mode = MissMode.CancelOnNextMiss)
            {
                MaxMissCount = maxMissCount;
                Mode = mode;
            }
        }

        public class SearchIgnoringInfo
        {
            public enum IgnoringMode
            {
                None,               // не игнорировать пробелы и точки
                IgnoreSpaces,
                IgnoreSpacesAndDots
            }

            public int? MaxIgnoreCount { get; set; }         // по сути - сколько слов возвращать
            public IgnoringMode Mode { get; set; }

            public SearchIgnoringInfo(int? maxIgnoreCount, IgnoringMode mode = IgnoringMode.None)
            {
                MaxIgnoreCount = maxIgnoreCount;
                Mode = mode;
            }
        }

        public class StringSearchResult
        {
            public string FoundString { get; set; }
            public int TextBreakIndex { get; set; }
            public int HtmlBreakIndex { get; set; }

            public bool MissFound { get; set; }

            public StringSearchResult()
            {
                FoundString = string.Empty;
            }
        }

        #endregion

        public string InputString { get; set; }
        public int StartIndex { get; set; }
        public string Alphabet { get; set; }
        public SearchDirection Direction { get; set; }
        public StringSearchMode SearchMode { get; set; }
        public SearchMissInfo MissInfo { get; set; }
        public SearchIgnoringInfo IgnoringInfo { get; set; }

        private StringSearchResult _searchResult;
        private int? _missCount = null;
        private int? _invalidSymbolsCount = null;   //  количество промахов. null - если ещё не начали находить значащие символы
        private bool _prevSymbolWasInvalid = false;
        private bool? _isDigits = null;  // true - ищем цифры, false - ищем текст 
        private int _indexBeforeHtml = -1;
        private int _index;

        public StringSearcher()
        {            
        }

        public StringSearcher(
            string s,
            int index,            
            SearchDirection searchDirection,            
            StringSearchMode searchMode = StringSearchMode.NotSpecified,
            SearchMissInfo missInfo = null,
            SearchIgnoringInfo ignoringInfo = null): this()
        {
            this.InputString = s;
            this.StartIndex = index;
            //this.Alphabet = alphabet;
            this.Direction = searchDirection;
            this.SearchMode = searchMode;
            this.MissInfo = missInfo;
            this.IgnoringInfo = ignoringInfo;
        }

        public static StringSearchResult SearchInString(
            string s,
            int index,
            SearchDirection searchDirection,
            StringSearchMode searchMode = StringSearchMode.NotSpecified,
            SearchMissInfo missInfo = null,
            SearchIgnoringInfo ignoringInfo = null)
        {
            return SearchInString(s, index, searchDirection, null, searchMode, missInfo, ignoringInfo);
        }

        public static StringSearchResult SearchInString(
            string s,
            int index,
            SearchDirection searchDirection,
            string alphabet,
            StringSearchMode searchMode = StringSearchMode.NotSpecified,
            SearchMissInfo missInfo = null,
            SearchIgnoringInfo ignoringInfo = null)
        {
            var searcher = new StringSearcher(s, index, searchDirection, searchMode, missInfo, ignoringInfo);
            searcher.Alphabet = alphabet;
            
            return searcher.SearchInString();
        }

        public StringSearchResult SearchInString()
        {
            SetDefaults();

            CheckArguments();

            if (SearchMode == StringSearchMode.SearchText)
                _isDigits = false;
            else if (SearchMode == StringSearchMode.SearchNumber)
                _isDigits = true;

            _searchResult = new StringSearchResult();  

            if (Direction == SearchDirection.Back)
            {
                for (_index = StartIndex - 1; _index >= 0; _index--)
                {
                    if (ProcessChar())
                        break;
                }
            }
            else
            {
                for (_index = StartIndex + 1; _index < InputString.Length; _index++)
                {
                    if (ProcessChar())
                        break;
                }
            }

            _searchResult.TextBreakIndex = _indexBeforeHtml != -1 ? _indexBeforeHtml : _index;
            _searchResult.HtmlBreakIndex = _index;
            return _searchResult;
        }

        private void SetDefaults()
        {
            if (IgnoringInfo == null)
                IgnoringInfo = new SearchIgnoringInfo(null);
            
            if (MissInfo == null)
                MissInfo = new SearchMissInfo(null);
        }

        private bool ProcessChar()
        {
            var c = InputString[_index];

            if (Direction == SearchDirection.Back && c == '>')
            {
                if (_indexBeforeHtml == -1)  // если мы уже не идём чисто по тэгам
                    _indexBeforeHtml = _index;
                _index = InputString.LastIndexOf('<', _index - 1);
                return false;
            }
            else if (Direction == SearchDirection.Forward && c == '<')
            {
                if (_indexBeforeHtml == -1)  // если мы уже не идём чисто по тэгам
                    _indexBeforeHtml = _index;
                _index = InputString.IndexOf('>', _index + 1);
                return false;
            }

            if (SearchMode == StringSearchMode.SearchFirstChar)
            {
                _searchResult.FoundString = c.ToString();
                return true;
            }

            bool toBreak;
            if (StringUtils.IsDigit(c))  // значит цифры
            {
                toBreak = ProcessDigitOrAlphabeticalChar(c, true);
            }
            else if (StringUtils.IsCharAlphabetical(c, Alphabet))
            {
                toBreak = ProcessDigitOrAlphabeticalChar(c, false);
            }
            else
            {
                toBreak = ProcessInvalidChar(c);
            }

            if (!toBreak)
                _indexBeforeHtml = -1;  // то есть сделали удачный круг, символ был нужным, значит нам нет смысла возвращаться вначало

            return toBreak;
        }      

        private bool ProcessDigitOrAlphabeticalChar(char c, bool isDigit)
        {
            if (!_invalidSymbolsCount.HasValue)
                _invalidSymbolsCount = 0;
            _prevSymbolWasInvalid = false;

            if (!_isDigits.HasValue)
            {
                _isDigits = isDigit;
                AddToResult(c);                
            }
            else
            {
                if (_isDigits.Value != isDigit)
                {
                    if (OnMissFound())
                        return true;
                }

                AddToResult(c);

                if (CheckIfCancelOnMissFound())
                    return true;
            }

            return false;
        }

        private bool ProcessInvalidChar(char c)
        {
            if (SearchMode == StringSearchMode.SearchFirstValueChar)
            {
                _searchResult.FoundString = c.ToString();
                return true;
            }
            else
            {
                var isMiss = true;
                if (_invalidSymbolsCount.HasValue && !_prevSymbolWasInvalid)
                    _invalidSymbolsCount++;
                _prevSymbolWasInvalid = true;

                if (c == ' ' || c == '.')
                {
                    if (_invalidSymbolsCount.GetValueOrDefault(0) < IgnoringInfo.MaxIgnoreCount.GetValueOrDefault(0))
                    {
                        switch (IgnoringInfo.Mode)
                        {
                            case SearchIgnoringInfo.IgnoringMode.IgnoreSpaces:
                                if (c == ' ')
                                    isMiss = false;
                                break;
                            case SearchIgnoringInfo.IgnoringMode.IgnoreSpacesAndDots:
                                isMiss = false;
                                break;
                        }
                    }
                }

                if (isMiss)
                {
                    if (OnMissFound())
                        return true;
                }

                AddToResult(c);

                if (CheckIfCancelOnMissFound())
                    return true;
            }

            return false;
        }

        private bool OnMissFound()
        {
            _searchResult.MissFound = true;
            _missCount = _missCount.GetValueOrDefault(0) + 1;
            if (_missCount > MissInfo.MaxMissCount.GetValueOrDefault(0))
                return true;

            return false;
        }

        private bool CheckIfCancelOnMissFound()
        {
            if (MissInfo.Mode == SearchMissInfo.MissMode.CancelOnMissFound && _missCount.GetValueOrDefault(-1) == MissInfo.MaxMissCount.GetValueOrDefault(0))  // -1 специально передаётся, чтобы отсечь варианты, когда missCount == null
            {
                _indexBeforeHtml = -1;  // мы взяли последний символ и он для нас был важным. значит и текст идёт до сюда, а не заканчивается там, где мы в прошлый раз начали какой нибудь тэг.

                if (Direction == SearchDirection.Back)
                    _index--;
                else
                    _index++;

                return true;
            }

            return false;
        }

        private void AddToResult(char c)
        {
            if (Direction == SearchDirection.Back)
                _searchResult.FoundString = c.ToString() + _searchResult.FoundString;
            else
                _searchResult.FoundString += c.ToString();
        }

        private void CheckArguments()
        {
            if (SearchMode == StringSearchMode.SearchFirstValueChar || SearchMode == StringSearchMode.SearchFirstChar)
            {
                if (IgnoringInfo.Mode != SearchIgnoringInfo.IgnoringMode.None)
                    throw new ArgumentException("Conflict parameters values: searchMode and ignoreSpaces");

                if (MissInfo.MaxMissCount.HasValue)
                    throw new ArgumentException("Conflict parameters values: searchMode and maxMissCount");
            }
        }
    }
}
