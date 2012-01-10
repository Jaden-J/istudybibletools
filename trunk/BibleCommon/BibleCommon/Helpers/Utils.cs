using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Xml;
using Microsoft.Office.Interop.OneNote;
using System.Reflection;
using System.IO;
using BibleCommon.Consts;
using BibleCommon.Services;
using System.Xml.XPath;

namespace BibleCommon.Helpers
{
    public static class Utils
    {
        public static string GetNotebookId(Application oneNoteApp, string notebookName)
        {
            string xml;
            XmlNamespaceManager xnm;
            oneNoteApp.GetHierarchy(null, HierarchyScope.hsNotebooks, out xml);
            XDocument doc = Utils.GetXDocument(xml, out xnm);
            XElement bibleNotebook = doc.Root.XPathSelectElement(string.Format("one:Notebook[@name='{0}']", notebookName), xnm);
            if (bibleNotebook != null)
            {
                return (string)bibleNotebook.Attribute("ID");
            }

            return null;
        }

        public static string GetHierarchyElementName(Application oneNoteApp, string elementId)
        {
            XmlNamespaceManager xnm;
            string xml = OneNoteProxy.Instance.GetHierarchy(oneNoteApp, elementId, HierarchyScope.hsSelf);
            XDocument doc = GetXDocument(xml, out xnm);
            return (string)doc.Root.Attribute("name");
        }

        public static XDocument GetXDocument(string xml, out XmlNamespaceManager xnm)
        {
            XDocument xd = XDocument.Parse(xml);
            xnm = new XmlNamespaceManager(new NameTable());
            xnm.AddNamespace("one", Constants.OneNoteXmlNs);
            return xd;
        }

        // возвращает количество родительских узлов
        public static int GetDepthLevel(XElement element)
        {
            int result = 0;

            if (element.Parent != null)
            {                
                result += 1 + GetDepthLevel(element.Parent);
            }

            return result;
        }

        public static bool IsSurroundedBy(string s, string leftSymbol, string rightSymbol, int startPosition = 0)
        {
            bool isSurroundedOnRight = false;
            bool isSurroundedOnLeft = false;

            int startIndex = s.IndexOf(leftSymbol, startPosition);
            int endIndex = s.IndexOf(rightSymbol, startPosition);
            if (!((startIndex == -1 && endIndex == -1) || (startIndex != -1 && endIndex == -1)
                || (startIndex != -1 && startIndex < endIndex)))                  // в любом случае здесь endIndex != -1, иначе бы он на предыдущем условии вышел                    
                isSurroundedOnRight = true;

            if (isSurroundedOnRight)
            {                
                startIndex = s.LastIndexOf(rightSymbol, startPosition);
                endIndex = s.LastIndexOf(leftSymbol, startPosition);
                if (!((startIndex == -1 && endIndex == -1) || (startIndex != -1 && endIndex == -1)
                    || (startIndex != -1 && startIndex > endIndex)))                  // в любом случае здесь endIndex != -1, иначе бы он на предыдущем условии вышел                    
                    isSurroundedOnLeft = true;
            }

            return isSurroundedOnLeft && isSurroundedOnRight;
        }

        public static string GenerateHref(Application oneNoteApp, string title, string pageId, string objectId)        
        {
            string link;
            oneNoteApp.GetHyperlinkToObject(pageId, objectId, out link);

            return string.Format("<a href=\"{0}\">{1}</a>", link, title);
        }

        public static bool IsDigit(char c)
        {
            return (c >= '0' && c <= '9');
        }

        public static bool IsCharAlphabetical(char c)
        {
            return ((c >= 'а' && c <= 'я')
                 || (c >= 'А' && c <= 'Я')
                 || (c == 'ё' || c == 'Ё')
                 || (c >= 'a' && c <= 'z')
                 || (c >= 'A' && c <= 'Z'));
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
        /// возвращает номер, находящийся в начале строки: например вернёт 12 для строки "12 глава"
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

        /// <summary>
        /// возвращает номер, находящийся в конце строки: например вернёт 12 для строки "глава 12"
        /// </summary>
        /// <param name="pointerElement"></param>
        /// <returns></returns>
        public static int? GetStringLastNumber(string s)
        {
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

                return int.Parse(d3 + d2 + d1);
            }

            return null;
        }

        public static string GetDigit(string s, int index)
        {
            int d;
            if (index > 0 && index < s.Length)
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

        public enum StringSearchMode
        {
            NotSpecified,
            SearchNumber,
            SearchText,
            SearchFirstValueChar,   // получить первый значащий символ-разделитель (не цифра, и не симфол алфавита, не равный '<' и '>')
            SearchFirstChar         // получить первый любой символ (не равный '<' и '>')
        }

        public enum StringSearchSpaceIgnorance
        {
            None,               // не игнорировать пробелы
            IgnoreFirstSpace,
            IgnoreAllSpaces
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
        
        /// <summary>
        /// 
        /// </summary>
        /// <param name="s"></param>
        /// <param name="index"></param>
        /// <param name="maxMissCount">максимальное количество "промахов" - допустимое количество ошибочных символов</param>
        /// <param name="breakIndex"></param>
        /// <param name="ignoreSpaces">Правила игнорирования проблелов. Но если режим searchMode установлен в SearchFirstValueChar, то данный параметр игнорируется, так как пробел тоже считается разделителем</param>
        /// <param name="searchMode">что ищем</param>
        /// <returns></returns>
        public static string GetPrevString(string s, int index, SearchMissInfo missInfo, out int breakIndex, 
            StringSearchSpaceIgnorance ignoreSpaces = StringSearchSpaceIgnorance.None, StringSearchMode searchMode = StringSearchMode.NotSpecified)
        {
            if (searchMode == StringSearchMode.SearchFirstValueChar || searchMode == StringSearchMode.SearchFirstChar)
            {
                if (ignoreSpaces != StringSearchSpaceIgnorance.None)
                    throw new ArgumentException("Противоречащие значения параметров searchMode и ignoreSpaces");

                if (missInfo != null)
                    if (missInfo.MaxMissCount.HasValue)
                        throw new ArgumentException("Противоречащие значения параметров searchMode и maxMissCount");                
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

            int i;
            for (i = index - 1; i >= 0; i--)
            {
                char c = s[i];

                if (c == '>')
                {
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
                            i--;
                            break;
                        }
                    }
                }
                else if (IsCharAlphabetical(c))
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

                        if (c == ' ')
                        {
                            switch (ignoreSpaces)
                            {
                                case StringSearchSpaceIgnorance.IgnoreAllSpaces:
                                    isMiss = false;
                                    break;
                                case StringSearchSpaceIgnorance.IgnoreFirstSpace:
                                    if (!foundValidChars)  // значит первый пробел
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
                            i--;
                            break;
                        }
                    }
                }
            }
            
            breakIndex = i;
            return result;
        }

        public static string GetNextString(string s, int index, SearchMissInfo missInfo, out int breakIndex, 
            StringSearchSpaceIgnorance ignoreSpaces = StringSearchSpaceIgnorance.None, StringSearchMode searchMode = StringSearchMode.NotSpecified)
        {
            if (searchMode == StringSearchMode.SearchFirstValueChar || searchMode == StringSearchMode.SearchFirstChar)
            {
                if (ignoreSpaces != StringSearchSpaceIgnorance.None)
                    throw new ArgumentException("Противоречащие значения параметров searchMode и ignoreSpaces");

                if (missInfo != null)
                    if (missInfo.MaxMissCount.HasValue)
                        throw new ArgumentException("Противоречащие значения параметров searchMode и maxMissCount");
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

            int i;
            for (i = index + 1; i < s.Length; i++)
            {
                char c = s[i];

                if (c == '<')
                {
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
                            i++;
                            break;
                        }
                    }
                }
                else if (IsCharAlphabetical(c))
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

                        if (c == ' ')
                        {
                            switch (ignoreSpaces)
                            {
                                case StringSearchSpaceIgnorance.IgnoreAllSpaces:
                                    isMiss = false;
                                    break;
                                case StringSearchSpaceIgnorance.IgnoreFirstSpace:
                                    if (!foundValidChars)  // значит первый пробел
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
                            i++;
                            break;
                        }
                    }
                }
            }

            breakIndex = i;
            return result;
        }

        public static string GetNextCloseTag(string s, int index)
        {
            int startIndex = s.IndexOf("<", index + 1);
            if (startIndex != -1)
            {
                if (GetChar(s, startIndex + 1) == '/')
                {
                    int endIndex = s.IndexOf(">", startIndex + 2);
                    if (endIndex != -1)
                    {
                        return s.Substring(startIndex + 2, endIndex - startIndex - 2);
                    }
                }
            }
            return string.Empty;
        }

        public static string GetCurrentDirectory()
        {
            var assembly = Assembly.GetExecutingAssembly().CodeBase;
            var uri = new Uri(assembly);
            return Path.GetDirectoryName(uri.LocalPath);
        }
    }
}
