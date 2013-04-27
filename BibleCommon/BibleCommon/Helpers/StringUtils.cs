﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BibleCommon.Services;
using System.Text.RegularExpressions;

namespace BibleCommon.Helpers
{
    public enum StringSearchMode
    {
        NotSpecified,
        SearchNumber,
        SearchText,
        SearchFirstValueChar,   // получить первый значащий символ-разделитель (не цифра, и не симфол алфавита, не равный '<' и '>')
        SearchFirstChar         // получить первый любой символ (не равный '<' и '>')
    }

    public enum StringSearchIgnorance
    {
        None,               // не игнорировать пробелы и точки
        IgnoreFirstSpaces,
        IgnoreFirstSpacesAndDots,
        IgnoreAllSpacesAndDots        

    }

    public class SearchMissInfo
    {
        public enum MissMode
        {
            CancelOnMissFound,      // как только нашли неподходящий символ - добавляем его к результативной строке и выходим
            CancelOnNextMiss        // выходим, когда количество промахов превысило лемит
        }

        public int? MaxMissCount { get; set; }
        public MissMode Mode { get; set; }

        public SearchMissInfo(int? maxMissCount, MissMode mode = MissMode.CancelOnNextMiss)
        {
            this.MaxMissCount = maxMissCount;
            this.Mode = mode;
        }
    }    

    public static class StringUtils
    {
        private static readonly Regex htmlPattern = new Regex(@"<(.|\n)*?>", RegexOptions.Compiled);      

        public static string GetText(string htmlString)
        {
            return htmlPattern.Replace(htmlString, string.Empty);
        }

        //public static string GetText(string htmlString, string alphabet)
        //{
        //    int t1, t2;
        //    string s = StringUtils.GetNextString(htmlString, -1,
        //                            new SearchMissInfo(htmlString.Length, SearchMissInfo.MissMode.CancelOnNextMiss), alphabet, out t1, out t2);

        //    return s;
        //}
        
        /// <summary>
        /// Сортирует, учитывая тот факт, что переданные строки могут содержать вначале цифры и сортировать надо по этим цифрам
        /// </summary>
        /// <param name="s1"></param>
        /// <param name="s2"></param>
        /// <returns></returns>
        public static int CompareTo(string s1, string s2)
        {
            if (!string.IsNullOrEmpty(s1))
            {
                if (string.IsNullOrEmpty(s2))
                    return 1;
                
                int temp;
                if (int.TryParse(s1[0].ToString(), out temp) && int.TryParse(s2[0].ToString(), out temp))
                {
                    int i1 = GetStringFirstNumber(s1).Value;
                    int i2 = GetStringFirstNumber(s2).Value;

                    return i1.CompareTo(i2);
                }
                else
                    return s1.CompareTo(s2);
            }

            return string.IsNullOrEmpty(s2) ? 0 : -1;
        }

        public static int IndexOfAny(string s, params string[] anyOf)
        {
            int minIndex = -1;

            foreach (string pattern in anyOf)
            {
                int i = s.IndexOf(pattern);
                if (i != -1)
                {
                    if (minIndex == -1 || i < minIndex)                    
                        minIndex = i;                    
                }                
            }

            return minIndex;

        }

        public static int LastIndexOf(string s, string value, int startIndex, int endIndex)
        {
            int result = -1;
            int i = s.IndexOf(value, startIndex);

            while (i > -1)
            {
                if (i <= endIndex)
                    result = i;
                else
                    break;

                i = s.IndexOf(value, i + 1);
            }
            
            return result;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="s"></param>
        /// <param name="leftSymbol"></param>
        /// <param name="rightSymbol"></param>
        /// <param name="startPosition"></param>
        /// <param name="searchInHtml"></param>
        /// <param name="textString">если мы ищем по html строке (searchInHtml == true), и если результат == true, то тогда здесь содержится текст между leftSymbol и rightSymbol</param>
        /// <returns></returns>
        public static bool IsSurroundedBy(string s, string leftSymbol, string rightSymbol, int startPosition, bool searchInHtml, out string textString)
        {
            textString = null;
            bool isSurroundedOnRight = false;
            bool isSurroundedOnLeft = false;                        
            string rightString = searchInHtml ? string.Empty : GetText(s.Substring(startPosition + 1));

            int htmlLeftIndex = 0, htmlRightIndex = 0;

            int startIndex = searchInHtml ? s.IndexOf(leftSymbol, startPosition) : rightString.IndexOf(leftSymbol);
            int endIndex = searchInHtml ? s.IndexOf(rightSymbol, startPosition) : rightString.IndexOf(rightSymbol);
            if (!((startIndex == -1 && endIndex == -1) || (startIndex != -1 && endIndex == -1)
                || (startIndex != -1 && startIndex < endIndex)))                  // в любом случае здесь endIndex != -1, иначе бы он на предыдущем условии вышел    
            {
                htmlRightIndex = endIndex;
                isSurroundedOnRight = true;
            }

            if (isSurroundedOnRight)
            {
                string leftString = searchInHtml ? string.Empty : GetText(s.Substring(0, startPosition));

                startIndex = searchInHtml ? s.LastIndexOf(rightSymbol, startPosition) : leftString.LastIndexOf(rightSymbol);
                endIndex = searchInHtml ? s.LastIndexOf(leftSymbol, startPosition) : leftString.LastIndexOf(leftSymbol);
                if (!((startIndex == -1 && endIndex == -1) || (startIndex != -1 && endIndex == -1)
                    || (startIndex != -1 && startIndex > endIndex)))                  // в любом случае здесь endIndex != -1, иначе бы он на предыдущем условии вышел                    
                {
                    htmlLeftIndex = endIndex;
                    isSurroundedOnLeft = true;
                }
            }

            var result = isSurroundedOnLeft && isSurroundedOnRight;
            if (result && searchInHtml)
                textString = s.Substring(htmlLeftIndex, htmlRightIndex - htmlLeftIndex + rightSymbol.Length + 1);

            return result;                
        }

        public static bool IsDigit(char c)
        {
            return char.IsDigit(c);            
        }

        public static bool IsCharAlphabetical(char c)
        {
            return IsCharAlphabetical(c, null);
        }

        public static bool IsCharAlphabetical(char c, string alphabet, bool strict = false)
        {
            if (strict && string.IsNullOrEmpty(alphabet))
            {
                if (!string.IsNullOrEmpty(SettingsManager.Instance.ModuleShortName))
                    alphabet = SettingsManager.Instance.CurrentModuleCached.BibleStructure.Alphabet;
            }

            if (strict && !string.IsNullOrEmpty(alphabet))            
                return alphabet.Contains(c);                                                            
            else
                return char.IsLetter(c);
        }

        public static int GetEntranceCount(string s, string searchString)
        {
            int result = 0;

            int i = s.IndexOf(searchString);

            while (i > -1)
            {
                result++;
                i = s.IndexOf(searchString, i + 1);
            }

            return result;
        }

        /// <summary>
        /// возвращает номер, находящийся первым в строке: например вернёт 12 для строки " 12 глава"
        /// ограничение: поддерживает максимум трёхзначные числа
        /// </summary>
        /// <param name="pointerElement"></param>
        /// <returns></returns>
        public static int? GetStringFirstNumber(string s, int startIndex = 0)
        {
            int i = s.IndexOfAny(new char[] { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9' }, startIndex);
            if (i != -1)
            {
                string d1 = s[i].ToString();
                string d2 = string.Empty;
                string d3 = string.Empty;

                d2 = GetDigit(s, i + 1);
                if (!string.IsNullOrEmpty(d2))
                    d3 = GetDigit(s, i + 2);

                return int.Parse(d1 + d2 + d3);
            }

            return null;
        }

        public static int? GetStringLastNumber(string s)
        {
            int index;
            return GetStringLastNumber(s, out index);
        }

        /// <summary>
        /// возвращает номер, находящийся последним в строке: например вернёт 12 для строки "глава 12 "
        /// </summary>
        /// <param name="pointerElement"></param>
        /// <returns></returns>
        public static int? GetStringLastNumber(string s, out int index)
        {
            index = -1;

            int i = s.LastIndexOfAny(new char[] { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9' });
            if (i != -1)
            {
                string d1 = s[i].ToString();
                string d2 = string.Empty;
                string d3 = string.Empty;

                if (i - 1 >= 0)
                    d2 = GetDigit(s, i - 1);

                if (!string.IsNullOrEmpty(d2))
                    if (i - 2 >= 0)
                        d3 = GetDigit(s, i - 2);

                index = i;
                return int.Parse(d3 + d2 + d1);
            }

            return null;
        }

        public static string GetDigit(string s, int index)
        {
            int d;
            if (index > -1 && index < s.Length)
                if (int.TryParse(s[index].ToString(), out d))
                    return d.ToString();

            return string.Empty;
        }

        public static char GetChar(string s, int index)
        {
            if (index >= 0 && index < s.Length)
                return s[index];

            return default(char);
        }

        public static string GetAttributeValue(string s, string attributeName)
        {
            string result = null;

            string searchString = string.Format("{0}=", attributeName);

            int startIndex = s.IndexOf(searchString);

            int symbolIndex = startIndex + searchString.Length;

            if (startIndex > -1 && s.Length > symbolIndex)
            {
                var c = s[symbolIndex];
                if (c == '{')
                    c = '}';
                else if (c == '(')
                    c = ')';
                else if (c == '[')
                    c = ']';

                int endIndex = s.IndexOf(c, symbolIndex + 1);

                if (endIndex > -1)
                {
                    result = s.Substring(symbolIndex + 1, endIndex - symbolIndex - 1);
                }
            }

            return result;
        }

        /// <summary>
        /// Получить значение аттрибута, не обрамлённого символами
        /// </summary>
        /// <param name="s"></param>
        /// <param name="attributeName"></param>
        /// <returns></returns>
        public static string GetNotFramedAttributeValue(string s, string attributeName)
        {
            string result = null;

            var searchString = string.Format("&{0}=", attributeName);

            var startIndex = s.IndexOf(searchString);

            if (startIndex > -1)
            {
                startIndex = startIndex + searchString.Length;

                if (s.Length > startIndex)
                {
                    int endIndex = s.IndexOf("&", startIndex + 1);

                    if (endIndex > -1)
                        result = s.Substring(startIndex, endIndex - startIndex);
                    else
                        result = s.Substring(startIndex);
                }
            }
            return result;
        }

        public static string GetQueryParameterValue(string s, string key)
        {
            string result = null;

            string searchString = string.Format("{0}=", key);

            int startIndex = s.IndexOf(searchString);

            if (startIndex > -1)
            {
                startIndex = startIndex + searchString.Length;

                var endIndex = s.IndexOfAny(new char[] { '&', '\"', '\'' }, startIndex + 1);

                if (endIndex == -1)
                    result = s.Substring(startIndex);
                else
                    result = s.Substring(startIndex, endIndex - startIndex);
            }

            return result;
        }        


        public static string MultiplyString(string s, int count)
        {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < count; i++)
            {
                sb.Append(s);
            }

            return sb.ToString();
        }


        public static string GetPrevString(string s, int index, SearchMissInfo missInfo, out int textBreakIndex, out int htmlBreakIndex,
           StringSearchIgnorance ignoreSpaces = StringSearchIgnorance.None, StringSearchMode searchMode = StringSearchMode.NotSpecified)
        {
            return GetPrevString(s, index, missInfo, null, out textBreakIndex, out htmlBreakIndex, ignoreSpaces, searchMode);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="s"></param>
        /// <param name="index"></param>
        /// <param name="maxMissCount">максимальное количество "промахов" - допустимое количество ошибочных символов</param>
        /// <param name="textBreakIndex"></param>
        /// <param name="ignoreSpaces">Правила игнорирования проблелов. Но если режим searchMode установлен в SearchFirstValueChar, то данный параметр игнорируется, так как пробел тоже считается разделителем</param>
        /// <param name="searchMode">что ищем</param>
        /// <returns></returns>
        public static string GetPrevString(string s, int index, SearchMissInfo missInfo, string alphabet, out int textBreakIndex, out int htmlBreakIndex,
            StringSearchIgnorance ignoreSpaces = StringSearchIgnorance.None, StringSearchMode searchMode = StringSearchMode.NotSpecified)
        {
            if (searchMode == StringSearchMode.SearchFirstValueChar || searchMode == StringSearchMode.SearchFirstChar)
            {
                if (ignoreSpaces != StringSearchIgnorance.None)
                    throw new ArgumentException("Conflict parameters values: searchMode and ignoreSpaces");

                if (missInfo != null)
                    if (missInfo.MaxMissCount.HasValue)
                        throw new ArgumentException("Conflict parameters values: searchMode and maxMissCount");
            }

            if (missInfo == null)
                missInfo = new SearchMissInfo(null);

            string result = string.Empty;
            int? missCount = null;

            bool foundValidChars = false;   // уже начали чот нить находить
            bool? isDigits = null;  // true - ищем цифры, false - ищем текст

            if (searchMode == StringSearchMode.SearchText)
                isDigits = false;
            else if (searchMode == StringSearchMode.SearchNumber)
                isDigits = true;

            int iBeforeHtml = -1;
            int i;
            for (i = index - 1; i >= 0; i--)
            {
                char c = s[i];

                if (c == '>')
                {
                    if (iBeforeHtml == -1)  // если мы уже не идём чисто по тэгам
                        iBeforeHtml = i;
                    i = s.LastIndexOf('<', i - 1);
                    continue;
                }

                if (searchMode == StringSearchMode.SearchFirstChar)
                {
                    result = c.ToString();
                    break;
                }

                if (IsDigit(c))  // значит цифры
                {
                    foundValidChars = true;

                    if (!isDigits.HasValue)
                    {
                        isDigits = true;
                        result = c.ToString() + result;
                    }
                    else
                    {
                        if (!isDigits.Value)
                        {
                            missCount = missCount.GetValueOrDefault(0) + 1;
                            if (missCount > missInfo.MaxMissCount.GetValueOrDefault(0))
                                break;
                        }

                        result = c.ToString() + result;

                        if (missInfo.Mode == SearchMissInfo.MissMode.CancelOnMissFound && missCount.GetValueOrDefault(-1) == missInfo.MaxMissCount.GetValueOrDefault(0))  // -1 специально передаётся, чтобы отсечь варианты, когда missCount == null
                        {
                            iBeforeHtml = -1;  // мы взяли последний символ и он для нас был важным. значит и текст идёт до сюда, а не заканчивается там, где мы в прошлый раз начали какой нибудь тэг.
                            i--;
                            break;
                        }
                    }
                }
                else if (IsCharAlphabetical(c, alphabet))
                {
                    foundValidChars = true;

                    if (!isDigits.HasValue)
                    {
                        isDigits = false;
                        result = c.ToString() + result;
                    }
                    else
                    {
                        if (isDigits.Value)
                        {
                            missCount = missCount.GetValueOrDefault(0) + 1;
                            if (missCount > missInfo.MaxMissCount.GetValueOrDefault(0))
                                break;
                        }

                        result = c.ToString() + result;

                        if (missInfo.Mode == SearchMissInfo.MissMode.CancelOnMissFound && missCount.GetValueOrDefault(-1) == missInfo.MaxMissCount.GetValueOrDefault(0))  // -1 специально передаётся, чтобы отсечь варианты, когда missCount == null
                        {
                            iBeforeHtml = -1;  // мы взяли последний символ и он для нас был важным. значит и текст идёт до сюда, а не заканчивается там, где мы в прошлый раз начали какой нибудь тэг.
                            i--;
                            break;
                        }
                    }
                }
                else
                {
                    if (searchMode == StringSearchMode.SearchFirstValueChar)
                    {
                        result = c.ToString();
                        break;
                    }
                    else
                    {
                        bool isMiss = true;

                        if (c == ' ' || c == '.')
                        {
                            switch (ignoreSpaces)
                            {
                                case StringSearchIgnorance.IgnoreAllSpacesAndDots:
                                    isMiss = false;
                                    break;
                                case StringSearchIgnorance.IgnoreFirstSpacesAndDots:
                                    if (!foundValidChars)  // значит пробел или точка до текста
                                        isMiss = false;
                                    break;
                                case StringSearchIgnorance.IgnoreFirstSpaces:
                                    if (c == ' ' && !foundValidChars)
                                        isMiss = false;
                                    break;
                            }
                        }

                        if (isMiss)
                        {
                            missCount = missCount.GetValueOrDefault(0) + 1;
                            if (missCount > missInfo.MaxMissCount.GetValueOrDefault(0))
                                break;
                        }

                        result = c.ToString() + result;

                        if (missInfo.Mode == SearchMissInfo.MissMode.CancelOnMissFound && missCount.GetValueOrDefault(-1) == missInfo.MaxMissCount.GetValueOrDefault(0))  // -1 специально передаётся, чтобы отсечь варианты, когда missCount == null
                        {
                            iBeforeHtml = -1;  // мы взяли последний символ и он для нас был важным. значит и текст идёт до сюда, а не заканчивается там, где мы в прошлый раз начали какой нибудь тэг.
                            i--;
                            break;
                        }
                    }
                }

                iBeforeHtml = -1;  // то есть сделали удачный круг, символ был нужным, значит нам нет смысла возвращаться вначало
            }

            textBreakIndex = iBeforeHtml != -1 ? iBeforeHtml : i;
            htmlBreakIndex = i;
            return result;
        }

        public static string GetNextString(string s, int index, SearchMissInfo missInfo, out int textBreakIndex, out int htmlBreakIndex,
            StringSearchIgnorance ignoreSpaces = StringSearchIgnorance.None, StringSearchMode searchMode = StringSearchMode.NotSpecified)
        {
            return GetNextString(s, index, missInfo, null, out textBreakIndex, out htmlBreakIndex, ignoreSpaces, searchMode);
        }


        public static string GetNextString(string s, int index, SearchMissInfo missInfo, string alphabet, out int textBreakIndex, out int htmlBreakIndex,
            StringSearchIgnorance ignoreSpaces = StringSearchIgnorance.None, StringSearchMode searchMode = StringSearchMode.NotSpecified)
        {
            if (searchMode == StringSearchMode.SearchFirstValueChar || searchMode == StringSearchMode.SearchFirstChar)
            {
                if (ignoreSpaces != StringSearchIgnorance.None)
                    throw new ArgumentException("Conflict parameters values: searchMode and ignoreSpaces");

                if (missInfo != null)
                    if (missInfo.MaxMissCount.HasValue)
                        throw new ArgumentException("Conflict parameters values: searchMode and maxMissCount");
            }

            if (missInfo == null)
                missInfo = new SearchMissInfo(null);

            string result = string.Empty;
            int? missCount = null;

            bool foundValidChars = false;   // уже начали чот нить находить
            bool? isDigits = null;  // true - ищем цифры, false - ищем текст

            if (searchMode == StringSearchMode.SearchText)
                isDigits = false;
            else if (searchMode == StringSearchMode.SearchNumber)
                isDigits = true;

            int iBeforeHtml = -1;
            int i;
            for (i = index + 1; i < s.Length; i++)
            {
                char c = s[i];

                if (c == '<')
                {
                    if (iBeforeHtml == -1)  // если мы уже не идём чисто по тэгам
                        iBeforeHtml = i;
                    i = s.IndexOf('>', i + 1);
                    continue;
                }

                if (IsDigit(c))  // значит цифры
                {
                    foundValidChars = true;

                    if (!isDigits.HasValue)
                    {
                        isDigits = true;
                        result += c.ToString();
                    }
                    else
                    {
                        if (!isDigits.Value)
                        {
                            missCount = missCount.GetValueOrDefault(0) + 1;
                            if (missCount > missInfo.MaxMissCount.GetValueOrDefault(0))
                                break;
                        }

                        result += c.ToString();

                        if (missInfo.Mode == SearchMissInfo.MissMode.CancelOnMissFound && missCount.GetValueOrDefault(-1) == missInfo.MaxMissCount.GetValueOrDefault(0))  // -1 специально передаётся, чтобы отсечь варианты, когда missCount == null
                        {
                            iBeforeHtml = -1;  // мы взяли последний символ и он для нас был важным. значит и текст идёт до сюда, а не заканчивается там, где мы в прошлый раз начали какой нибудь тэг.
                            i++;
                            break;
                        }
                    }
                }
                else if (IsCharAlphabetical(c, alphabet))
                {
                    foundValidChars = true;

                    if (!isDigits.HasValue)
                    {
                        isDigits = false;
                        result += c.ToString();
                    }
                    else
                    {
                        if (isDigits.Value)
                        {
                            missCount = missCount.GetValueOrDefault(0) + 1;
                            if (missCount > missInfo.MaxMissCount.GetValueOrDefault(0))
                                break;
                        }

                        result += c.ToString();

                        if (missInfo.Mode == SearchMissInfo.MissMode.CancelOnMissFound && missCount.GetValueOrDefault(-1) == missInfo.MaxMissCount.GetValueOrDefault(0))  // -1 специально передаётся, чтобы отсечь варианты, когда missCount == null
                        {
                            iBeforeHtml = -1;  // мы взяли последний символ и он для нас был важным. значит и текст идёт до сюда, а не заканчивается там, где мы в прошлый раз начали какой нибудь тэг.
                            i++;
                            break;
                        }
                    }
                }
                else
                {
                    if (searchMode == StringSearchMode.SearchFirstValueChar)
                    {
                        result = c.ToString();
                        break;
                    }
                    else
                    {
                        bool isMiss = true;

                        if (c == ' ' || c == '.')
                        {
                            switch (ignoreSpaces)
                            {
                                case StringSearchIgnorance.IgnoreAllSpacesAndDots:
                                    isMiss = false;
                                    break;
                                case StringSearchIgnorance.IgnoreFirstSpacesAndDots:
                                    if (!foundValidChars)  // значит пробел или точка до текста
                                        isMiss = false;
                                    break;
                                case StringSearchIgnorance.IgnoreFirstSpaces:
                                    if (c == ' ' && !foundValidChars)
                                        isMiss = false;
                                    break;
                            }
                        }

                        if (isMiss)
                        {
                            missCount = missCount.GetValueOrDefault(0) + 1;
                            if (missCount > missInfo.MaxMissCount.GetValueOrDefault(0))
                                break;
                        }

                        result += c.ToString();

                        if (missInfo.Mode == SearchMissInfo.MissMode.CancelOnMissFound && missCount.GetValueOrDefault(-1) == missInfo.MaxMissCount.GetValueOrDefault(0))  // -1 специально передаётся, чтобы отсечь варианты, когда missCount == null
                        {
                            iBeforeHtml = -1;  // мы взяли последний символ и он для нас был важным. значит и текст идёт до сюда, а не заканчивается там, где мы в прошлый раз начали какой нибудь тэг.
                            i++;
                            break;
                        }
                    }
                }

                iBeforeHtml = -1;  // то есть сделали удачный круг, символ был нужным, значит нам нет смысла возвращаться вначало
            }

            textBreakIndex = iBeforeHtml != -1 ? iBeforeHtml : i;
            htmlBreakIndex = i;
            return result;
        }

        public static int GetNextIndexOfDigit(string s, int? index)
        {
            if (s.Length >= index.GetValueOrDefault(0) + 2)
            {
                index = s
                    .IndexOfAny(new char[] { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9', '<' }, index.HasValue ? index.Value + 1 : 0);

                if (index != -1 && s[index.Value] == '<')
                {
                    index = s.IndexOf('>', index.Value + 1);
                    return GetNextIndexOfDigit(s, index);
                }
            }
            else
                index = -1;

            return index.Value;
        }

        //public static string GetNextCloseTag(string s, int index)
        //{
        //    int startIndex = s.IndexOf("<", index + 1);
        //    if (startIndex != -1)
        //    {
        //        if (GetChar(s, startIndex + 1) == '/')
        //        {
        //            int endIndex = s.IndexOf(">", startIndex + 2);
        //            if (endIndex != -1)
        //            {
        //                return s.Substring(startIndex + 2, endIndex - startIndex - 2);
        //            }
        //        }
        //    }
        //    return string.Empty;
        //}

        //public static string RemoveAllEntries(string s, IEnumerable<string> entries)
        //{
        //    foreach (var entry in entries)
        //        s = s.Replace(entry, string.Empty);

        //    return s;
        //}

        public static string RemoveTags(string s, string startString, string endString)
        {
            if (!string.IsNullOrEmpty(s))
            {
                int startIndex = s.IndexOf(startString);
                if (startIndex != -1)
                {
                    int endIndex = s.IndexOf(endString, startIndex);
                    if (startIndex != -1)
                        s = s.Substring(0, startIndex) + RemoveTags(s.Substring(endIndex + endString.Length), startString, endString);
                }
            }

            return s;
        }

        public static string RemoveIllegalTagStartAndEndSymbols(string s, int startIndex = 0)
        {
            startIndex = s.IndexOfAny(new char[] { '<', '>' }, startIndex);
            if (startIndex != -1)
            {
                if (s[startIndex] == '>')
                    return RemoveIllegalTagStartAndEndSymbols(s.Remove(startIndex, 1));
                else if (s.Length == startIndex + 1)
                    return s.Remove(startIndex, 1);
                else
                {
                    var nextChar = char.ToLower(s[startIndex + 1]);
                    if (!((nextChar >= 'a' && nextChar <= 'z') || nextChar == '/'))
                        return RemoveIllegalTagStartAndEndSymbols(s.Remove(startIndex, 1));
                    else
                        startIndex = s.IndexOf('>', startIndex + 1);
                }

                if (startIndex != -1)
                    return RemoveIllegalTagStartAndEndSymbols(s, startIndex + 1);
            }

            return s;
        }

        public static string ReplaceIgnoreCase(this string original,
                  string pattern, string replacement)
        {
            int count, position0, position1;
            count = position0 = position1 = 0;
            string upperString = original.ToUpper();
            string upperPattern = pattern.ToUpper();
            int inc = (original.Length / pattern.Length) *
                      (replacement.Length - pattern.Length);
            char[] chars = new char[original.Length + Math.Max(0, inc)];
            while ((position1 = upperString.IndexOf(upperPattern,
                                              position0)) != -1)
            {
                for (int i = position0; i < position1; ++i)
                    chars[count++] = original[i];
                for (int i = 0; i < replacement.Length; ++i)
                    chars[count++] = replacement[i];
                position0 = position1 + pattern.Length;
            }
            if (position0 == 0) return original;
            for (int i = position0; i < original.Length; ++i)
                chars[count++] = original[i];
            return new string(chars, 0, count);
        }

        /// <summary>
        /// Делает первую букву заглавной. Для строки "1кор" вернёт "1Кор"
        /// </summary>
        /// <param name="bookName"></param>
        /// <returns></returns>
        public static string MakeFirstLetterUpper(string bookName)
        {   
            if (!string.IsNullOrEmpty(bookName))
            {
                for (var i = 0; i < bookName.Length; i++)
                {
                    if (char.IsLetter(bookName[i]))                    
                        return string.Concat(bookName.Substring(0, i), char.ToUpperInvariant(bookName[i]), bookName.Substring(i + 1));                                                 
                }
            }

            return bookName;
        }
    }
}
