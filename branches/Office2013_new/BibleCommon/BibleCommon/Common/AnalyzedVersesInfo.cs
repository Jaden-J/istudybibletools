using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;
using System.ComponentModel;
using Polenter.Serialization;

namespace BibleCommon.Common
{
    [Serializable]
    [XmlRoot]
    public class AnalyzedVersesInfo
    {
        private const string StringDelimiter = "; ";

        private Dictionary<int, AnalyzedBookInfo> _booksDictionary;        
        private Dictionary<int, AnalyzedBookInfo> BooksDictionary
        {
            get
            {
                if (_booksDictionary == null)
                {
                    _booksDictionary = new Dictionary<int, AnalyzedBookInfo>();
                    Books.ForEach(b => _booksDictionary.Add(b.Index, b));
                }

                return _booksDictionary;
            }
        }

        [XmlArray("Notebooks")]
        [XmlArrayItem(typeof(AnalyzedNotebookInfo), ElementName = "Notebook")]
        [EditorBrowsable(EditorBrowsableState.Never)]
        public List<AnalyzedNotebookInfo> Notebooks {
            get
            {
                return NotebooksDictionary.Values.ToList();
            }
            set
            {
                NotebooksDictionary = new Dictionary<string, AnalyzedNotebookInfo>();
                value.ForEach(n => NotebooksDictionary.Add(n.Name, n));
            }
        }

        [XmlIgnore]
        public Dictionary<string, AnalyzedNotebookInfo> NotebooksDictionary { get; set; }

        [XmlArray("Books")]        
        [XmlArrayItem(typeof(AnalyzedBookInfo), ElementName="Book")]
        [EditorBrowsable(EditorBrowsableState.Never)]
        public List<AnalyzedBookInfo> Books { get; set; }        

        [XmlAttribute]
        public string Module { get; set; }

        public AnalyzedVersesInfo()
        {
            Books = new List<AnalyzedBookInfo>();
            Notebooks = new List<AnalyzedNotebookInfo>();
        }

        public AnalyzedVersesInfo(string module)
            : this()
        {
            Module = module;
        }

        public AnalyzedBookInfo GetOrCreateBookInfo(int bookIndex, string bookName)
        {
            if (!BooksDictionary.ContainsKey(bookIndex))
            {
                var bookInfo = new AnalyzedBookInfo() { Index = bookIndex, Name = bookName };
                BooksDictionary.Add(bookIndex, bookInfo);
                Books.Add(bookInfo);
                return bookInfo;
            }
            else
                return BooksDictionary[bookIndex];
        }

        public void Sort()
        {
            Books = Books.OrderBy(b => b.Index).ToList();
            Books.ForEach(b => b.Sort());
        }
    }

    public class AnalyzedNotebookInfo
    {
        [XmlAttribute]
        public string Name { get; set; }

        [XmlAttribute]
        public string Nickname { get; set; }

        [XmlAttribute]
        public int Id { get; set; }

        public override int GetHashCode()
        {
            return Name.GetHashCode();
        }

        public override bool Equals(object obj)
        {
            var other = (AnalyzedNotebookInfo)obj;

            if (obj == null)
                return false;

            return Name.Equals(other.Name);
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
                    Chapters.ForEach(c => _chaptersDictionary.Add(c.Index, c));
                }

                return _chaptersDictionary;
            }
        }

        [XmlAttribute]
        public int Index { get; set; }

        [XmlAttribute]
        public string Name { get; set; }

        [XmlElement(typeof(AnalyzedChapterInfo), ElementName = "Chapter")]
        [EditorBrowsable(EditorBrowsableState.Never)]
        public List<AnalyzedChapterInfo> Chapters { get; set; }

        public AnalyzedBookInfo()
        {
            Chapters = new List<AnalyzedChapterInfo>();
        }

        public AnalyzedChapterInfo GetOrCreateChapterInfo(int chapterIndex)
        {
            if (!ChaptersDictionary.ContainsKey(chapterIndex))
            {
                var chapterInfo = new AnalyzedChapterInfo() { Index = chapterIndex };
                ChaptersDictionary.Add(chapterIndex, chapterInfo);
                Chapters.Add(chapterInfo);
                return chapterInfo;
            }
            else
                return ChaptersDictionary[chapterIndex];
        }

        public void Sort()
        {
            Chapters = Chapters.OrderBy(c => c.Index).ToList();
            Chapters.ForEach(c => c.Sort());
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
                    Verses.ForEach(v => _versesDictionary.Add(v.Index, v));
                }

                return _versesDictionary;
            }
        }

        [XmlAttribute]
        public int Index { get; set; }

        [XmlElement(typeof(AnalyzedVerseInfo), ElementName = "Verse")]
        [EditorBrowsable(EditorBrowsableState.Never)]
        public List<AnalyzedVerseInfo> Verses { get; set; }

        public AnalyzedChapterInfo()
        {
            Verses = new List<AnalyzedVerseInfo>();
        }

        public AnalyzedVerseInfo GetOrCreateVerseInfo(int verseIndex)
        {
            if (!VersesDictionary.ContainsKey(verseIndex))
            {
                var verseInfo = new AnalyzedVerseInfo() { Index = verseIndex };
                VersesDictionary.Add(verseIndex, verseInfo);
                Verses.Add(verseInfo);
                return verseInfo;
            }
            else
                return VersesDictionary[verseIndex];
        }

        public void Sort()
        {
            Verses = Verses.OrderBy(v => v.Index).ToList();
        }
    }

    [Serializable]
    public class AnalyzedVerseInfo
    {
        public AnalyzedVerseInfo()
        {
            Notebooks = new HashSet<int>();
            IsDetailedOnly = true;
        }

        [XmlAttribute]
        public int Index { get; set; }

        [XmlAttribute]
        public decimal MaxWeight { get; set; }

        [XmlAttribute]
        public decimal MaxDetailedWeight { get; set; }

        [XmlAttribute]
        public bool IsDetailedOnly { get; set; }

        [XmlAttribute("Notebooks")]
        [EditorBrowsable(EditorBrowsableState.Never)]
        public string XmlNotebooks
        {
            get
            {
                return string.Join(",", Notebooks.OrderBy(n => n));
            }
            set
            {
                Notebooks = new HashSet<int>(value
                            .Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries)
                            .ToList()
                            .ConvertAll(s => int.Parse(s)));
            }
        }

        [XmlIgnore]
        [ExcludeFromSerialization]
        public HashSet<int> Notebooks { get; set; }

        //[XmlIgnore]
        //public int? PartIndex { get; set; }

        //[XmlAttribute("PartIndex")]
        //[EditorBrowsable(EditorBrowsableState.Never)]
        //public int XmlPartIndex
        //{
        //    get
        //    {
        //        return PartIndex.Value;
        //    }
        //    set
        //    {
        //        PartIndex = value;
        //    }
        //}        
    }
}
