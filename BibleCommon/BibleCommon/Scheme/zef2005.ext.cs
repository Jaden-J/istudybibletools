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

        public string GetVerseContent(SimpleVersePointer versPointer, out VerseNumber verseNumber, out bool isEmpty)
        {
            isEmpty = false;

            if (this.Chapters.Count < versPointer.Chapter)
                throw new ParallelChapterNotFoundException(versPointer, BaseVersePointerException.Severity.Warning);

            var chapter = this.Chapters[versPointer.Chapter - 1];

            var verseContent = chapter.GetVerse(versPointer.Verse);
            if (verseContent == null)
                throw new ParallelVerseNotFoundException(versPointer, BaseVersePointerException.Severity.Warning);

            verseNumber = verseContent.VerseNumber;

            if (verseContent.IsEmpty)
            {
                isEmpty = true;
                return string.Empty;
            }

            string result = null;

            if (versPointer.PartIndex.HasValue)
            {
                var versesParts = verseContent.Value.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                if (versesParts.Length > versPointer.PartIndex.Value)
                    result = versesParts[versPointer.PartIndex.Value].Trim();
            }
            else
                result = verseContent.Value;

            result = ShellVerseText(result);

            return result;
        }

        public string GetVersesContent(List<SimpleVersePointer> verses, out int? topVerse, out bool isEmpty)
        {
            var contents = new List<string>();

            topVerse = verses.First().TopVerse;

            isEmpty = true;

            foreach (var verse in verses)
            {
                bool localIsEmpty;
                VerseNumber vn;
                contents.Add(GetVerseContent(verse, out vn, out localIsEmpty));
                isEmpty = isEmpty && localIsEmpty;

                if (vn.TopVerse.GetValueOrDefault(-2) > topVerse.GetValueOrDefault(-1))
                    topVerse = vn.TopVerse;
            }

            if (!topVerse.HasValue && verses.Count > 1)
                topVerse = verses.Last().Verse;

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
                _versesDictionary.Add(verse.Index, verse);
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

        [XmlIgnore]
        public string Value
        {
            get
            {   
                return string.Concat(this.Items.Where(
                    item => 
                        item is GRAM
                     || item is STYLE
                     || item is SUP
                     || item is string).ToArray());
            }
        }
    }

    public partial class gr : GRAM
    {
    }

    public partial class GRAM
    {
        public override string ToString()
        {
            if (Text != null)
                return string.Concat(Text);

            return null;
        }
    }

    public partial class STYLE
    {
        public override string ToString()
        {
            if (Text != null)
                return string.Concat(Text);

            return null;
        }
    }

    public partial class SUP
    {
        public override string ToString()
        {
            if (Text != null)
                return string.Concat(Text);

            return null;
        }
    } 
}
