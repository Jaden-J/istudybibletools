using System.Linq;
using System.Collections.Generic;
using BibleCommon.Common;
using System;
using System.Xml.Serialization;
using BibleCommon.Helpers;

namespace BibleCommon.Scheme
{
    public partial class XMLBIBLE
    {
        [XmlIgnore]
        public List<BIBLEBOOK> Books
        {
            get
            {
                if (this.BIBLEBOOK != null)
                    return this.BIBLEBOOK.ToList();

                return new List<BIBLEBOOK>();
            }
        }       
    }

    public partial class BIBLEBOOK
    {
        [XmlIgnore]
        public int Index
        {
            get
            {
                return !string.IsNullOrEmpty(this.bnumber) ? int.Parse(this.bnumber) : default(int);
            }
        }

        [XmlIgnore]
        public List<CHAPTER> Chapters
        {
            get
            {
                if (this.Items != null)
                    return this.Items.ToList();

                return new List<CHAPTER>();
            }
        }

        public string GetVerseContent(SimpleVersePointer versePointer, string strongPrefix, 
            out VerseNumber verseNumber, out bool isEmpty, out bool isFullVerse)
        {
            isFullVerse = true;
            isEmpty = false;            

            if (versePointer.IsEmpty)
            {
                isEmpty = true;
                verseNumber = versePointer.VerseNumber;
                return null;
            }

            if (this.Chapters.Count < versePointer.Chapter)
                throw new ParallelChapterNotFoundException(versePointer, BaseVersePointerException.Severity.Warning);

            var chapter = this.Chapters[versePointer.Chapter - 1];

            var verse = chapter.GetVerse(versePointer.Verse);
            if (verse == null)
                throw new ParallelVerseNotFoundException(versePointer, BaseVersePointerException.Severity.Warning);

            verseNumber = verse.VerseNumber;

            if (verse.IsEmpty)
            {
                isEmpty = true;
                return string.Empty;
            }

            string result = null;

            var verseContent = verse.GetValue(true, strongPrefix);
            var shelledVerseContent = ShellVerseText(verseContent);

            if (versePointer.PartIndex.HasValue)
            {
                var versesParts = verseContent.Split(new char[] { '|' });
                if (versesParts.Length > versePointer.PartIndex.Value)
                    result = versesParts[versePointer.PartIndex.Value].Trim();

                result = ShellVerseText(result);
                if (result != shelledVerseContent)
                    isFullVerse = false;
            }
            else            
                result = shelledVerseContent;

            return result;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="verses"></param>
        /// <param name="strongPrefix"></param>
        /// <param name="topVerse"></param>
        /// <param name="isEmpty"></param>
        /// <param name="isFullVerses">Запрашиваемые стихи являются полными. А то стих может быть "Текст стиха|". То есть вроде как две части стиха, но первая часть равна всему стиху.</param>
        /// <param name="isDiscontinuous">Прерывистые стихи. Например 8,22. </param>
        /// <param name="notFoundVerses"></param>
        /// <returns></returns>
        public string GetVersesContent(List<SimpleVersePointer> verses, string strongPrefix, 
            out int? topVerse, out bool isEmpty, out bool isFullVerses, out bool isDiscontinuous, out List<SimpleVersePointer> notFoundVerses)
        {
            var contents = new List<string>();
            notFoundVerses = new List<SimpleVersePointer>();

            var firstVerse = verses.First();
            topVerse = firstVerse.TopVerse.GetValueOrDefault(firstVerse.Verse);

            isEmpty = true;
            isFullVerses = true;
            isDiscontinuous = false;                        

            foreach (var verse in verses)
            {
                bool localIsEmpty, localIsFullVerse;
                VerseNumber vn;
                var verseContent = GetVerseContent(verse, strongPrefix, out vn, out localIsEmpty, out localIsFullVerse);
                contents.Add(verseContent);

                if (!localIsEmpty)
                {
                    if (verseContent == null)
                        notFoundVerses.Add(verse);
                    else if (verseContent == string.Empty)
                        localIsEmpty = true;
                }

                isEmpty = isEmpty && localIsEmpty;
                isFullVerses = isFullVerses && localIsFullVerse;

                if (vn.Verse > topVerse + 1)
                    isDiscontinuous = true;

                if (vn.TopVerse.GetValueOrDefault(vn.Verse) > topVerse)
                    topVerse = vn.TopVerse.GetValueOrDefault(vn.Verse);                                
            }

            if (topVerse == firstVerse.Verse)
                topVerse = null;

            if (contents.All(c => c == null))
                return null;
            else
                return string.Join(" ", contents.ToArray());
        }

        public static string GetFullVerseString(int verseNumber, int? topVerseNumber, string verseText)
        {
            string verseNumberString = topVerseNumber.HasValue ? string.Format("{0}-{1}", verseNumber, topVerseNumber) : verseNumber.ToString();
            return string.Format("{0}<span> </span>{1}", verseNumberString, ShellVerseText(verseText));
        }

        private static string ShellVerseText(string verseText)
        {
            if (!string.IsNullOrEmpty(verseText))
                verseText = verseText.Replace("|", string.Empty);

            return verseText;
        }
    }

    public partial class CHAPTER
    {
        [XmlIgnore]
        public int Index
        {
            get
            {
                return !string.IsNullOrEmpty(this.cnumber) ? int.Parse(this.cnumber) : default(int);
            }
        }

        [XmlIgnore]
        public List<VERS> Verses
        {
            get
            {
                if (this.Items != null)
                    return this.Items.OfType<VERS>().ToList();

                return new List<VERS>();
            }
        }       

        private Dictionary<int, VERS> _versesDictionary;
        public VERS GetVerse(int verseNumber)
        {
            if (_versesDictionary == null)
                LoadVersesDictionary();

            if (_versesDictionary.ContainsKey(verseNumber))
                return _versesDictionary[verseNumber];
            else if (Verses.Any(v => v.IsMultiVerse && (v.Index <= verseNumber && verseNumber <= v.TopIndex)))
                return VERS.Empty;

            return null;
        }

        private void LoadVersesDictionary()
        {
            _versesDictionary = new Dictionary<int, VERS>();
            foreach (var verse in Verses)
            {
                if (!_versesDictionary.ContainsKey(verse.Index))
                    _versesDictionary.Add(verse.Index, verse);
            }
        }
    }

    public partial class VERS
    {
        [XmlIgnore]
        public int Index
        {
            get
            {
                return !string.IsNullOrEmpty(this.vnumber) ? int.Parse(this.vnumber) : default(int);                
            }
        }

        [XmlIgnore]
        public int? TopIndex
        {
            get
            {
                return !string.IsNullOrEmpty(this.e) ? (int?)int.Parse(this.e) : null;
            }
        }

        [XmlIgnore]
        public bool IsMultiVerse 
        {
            get 
            {
                return TopIndex.HasValue; 
            }
        }

        [XmlIgnore]
        public VerseNumber VerseNumber
        {
            get
            {
                return new VerseNumber(Index, TopIndex);
            }
        }

        [XmlIgnore]
        public static VERS Empty
        {
            get
            {                
                return new VERS();
            }
        }

        [XmlIgnore]
        public bool IsEmpty
        {
            get
            {
                return string.IsNullOrEmpty(this.vnumber);
            }
        }

        public string GetValue(bool includeStrongNumbers, string strongPrefix = null)
        {
            return GetVerseText(this.Items, includeStrongNumbers, strongPrefix);
        }

        [XmlIgnore]
        public string Value
        {
            get
            {
                return GetValue(false);
            }
        }

        public static string GetVerseText(object[] items, bool includeStrongNumbers = false, string strongPrefix = null)
        {
            if (items == null)
                return null;
 
            return string.Concat(items.Where(
                                    item =>
                                        item is GRAM
                                     || item is STYLE
                                     || item is SUP
                                     || item is string)
                                 .Select(
                                    item =>
                                    {
                                        if (item is GRAM && includeStrongNumbers)                                        
                                            return string.Format("{0} {1}", item.ToString(), ((GRAM)item).GetStrongNumbersString(strongPrefix));                                        
                                        else
                                            return item;
                                    }).ToArray())
                    .Trim()
                    .Replace("  ", " ");
        }
    }

    public partial class gr : GRAM
    {
    }

    public partial class GRAM
    {
        public override string ToString()
        {
            if (Items != null)
                return string.Concat(Items);

            return string.Empty;
        }

        public string GetStrongNumbersString(string strongPrefix)
        {
            if (string.IsNullOrEmpty(strongPrefix))
                return str;

            var strongNumbers = str.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            return string.Join(" ", strongNumbers.Select(sn => 
                string.Concat(strongPrefix, 
                                int.Parse(sn).ToString("0000"))
                ).ToArray());
        }
    }

    public partial class STYLE
    {
        public override string ToString()
        {
            if (Items != null)
                return string.Concat(Items);            

            return string.Empty;
        }
    }

    public partial class SUP
    {
        public override string ToString()
        {
            if (Items != null)
                return string.Concat(Items);            

            return string.Empty;
        }
    } 
}
