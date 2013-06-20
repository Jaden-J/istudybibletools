using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;
using System.ComponentModel;

namespace BibleCommon.Common
{
    [Serializable]
    [XmlRoot]
    public class AnalyzedVersesInfo
    {
        private Dictionary<int, AnalyzedBookInfo> _booksDictionary;        
        private Dictionary<int, AnalyzedBookInfo> BooksDictionary
        {
            get
            {
                if (_booksDictionary == null)
                {
                    _booksDictionary = new Dictionary<int, AnalyzedBookInfo>();
                    Books.ForEach(b => _booksDictionary.Add(b.BookIndex, b));
                }

                return _booksDictionary;
            }
        }

        [XmlElement(typeof(AnalyzedBookInfo), ElementName = "Book")]
        [EditorBrowsable(EditorBrowsableState.Never)]
        public List<AnalyzedBookInfo> Books { get; set; }      

        public AnalyzedBookInfo GetOrCreateBookInfo(int bookIndex)
        {
            if (!BooksDictionary.ContainsKey(bookIndex))
            {
                var bookInfo = new AnalyzedBookInfo();
                BooksDictionary.Add(bookIndex, bookInfo);
                Books.Add(bookInfo);
                return bookInfo;
            }
            else
                return BooksDictionary[bookIndex];
        }

        public AnalyzedVersesInfo()
        {
            Books = new List<AnalyzedBookInfo>();
        }
    }

    [Serializable]
    public class AnalyzedBookInfo
    {
        private Dictionary<int, AnalyzedChapterInfo> _chaptersDictionary;
        private Dictionary<int, AnalyzedChapterInfo> ChaptersDictionary
        {
            get
            {
                if (_chaptersDictionary == null)
                {
                    _chaptersDictionary = new Dictionary<int, AnalyzedChapterInfo>();
                    Chapters.ForEach(c => _chaptersDictionary.Add(c.ChapterIndex, c));
                }

                return _chaptersDictionary;
            }
        }

        [XmlAttribute]
        public int BookIndex { get; set; }

        [XmlAttribute]
        public string BookName { get; set; }

        [XmlElement(typeof(AnalyzedChapterInfo), ElementName = "Chapter")]
        [EditorBrowsable(EditorBrowsableState.Never)]
        public List<AnalyzedChapterInfo> Chapters { get; set; }

        public AnalyzedChapterInfo GetOrCreateChapterInfo(int chapterIndex)
        {
            if (!ChaptersDictionary.ContainsKey(chapterIndex))
            {
                var chapterInfo = new AnalyzedChapterInfo();
                ChaptersDictionary.Add(chapterIndex, chapterInfo);
                Chapters.Add(chapterInfo);
                return chapterInfo;
            }
            else
                return ChaptersDictionary[chapterIndex];
        }

        public AnalyzedBookInfo()
        {
            Chapters = new List<AnalyzedChapterInfo>();
        }
    }

    [Serializable]
    public class AnalyzedChapterInfo
    {
        private Dictionary<int, AnalyzedVerseInfo> _versesDictionary;
        private Dictionary<int, AnalyzedVerseInfo> VersesDictionary
        {
            get
            {
                if (_versesDictionary == null)
                {
                    _versesDictionary = new Dictionary<int, AnalyzedVerseInfo>();
                    Verses.ForEach(v => _versesDictionary.Add(v.VerseIndex, v));  а здесь не надо по PartIndex-у ещё различать?
                }

                return _versesDictionary;
            }
        }

        [XmlAttribute]
        public int ChapterIndex { get; set; }

        [XmlElement(typeof(AnalyzedVerseInfo), ElementName = "Verse")]
        [EditorBrowsable(EditorBrowsableState.Never)]
        public List<AnalyzedVerseInfo> Verses { get; set; }

        public AnalyzedChapterInfo()
        {
            Verses = new List<AnalyzedVerseInfo>();
        }
    }

    [Serializable]
    public class AnalyzedVerseInfo
    {
        [XmlAttribute]
        public int VerseIndex { get; set; }

        [XmlIgnore]
        public int? PartIndex { get; set; }

        [XmlAttribute("PartIndex")]
        [EditorBrowsable(EditorBrowsableState.Never)]
        public int XmlPartIndex
        {
            get
            {
                return PartIndex.Value;
            }
            set
            {
                PartIndex = value;
            }
        }

        [XmlAttribute]
        public decimal MaxWeigth { get; set; }
    }
}
