﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;
using BibleCommon.Helpers;
using System.ComponentModel;

namespace BibleCommon.Common
{
    public enum ContainerType
    {
        Single,
        Bible,
        BibleComments,
        BibleNotesPages,
        BibleStudy,
        NewTestament,
        OldTestament,
        Other
    }

    public enum ModuleType
    {
        Bible = 0,
        Strong = 1,
        Dictionary = 2        
    }

    [Serializable]
    [XmlRoot(ElementName = "IStudyBibleTools_Module")]
    public class ModuleInfo
    {
        [XmlAttribute]
        [DefaultValue((int)ModuleType.Bible)]
        public ModuleType Type { get; set; }

        [XmlAttribute]
        public string Version { get; set; }

        private string _moduleShortName;
        [XmlAttribute]
        public string ShortName
        {
            get
            {
                return !string.IsNullOrEmpty(_moduleShortName) ? _moduleShortName.ToLower() : _moduleShortName;
            }
            set
            {
                _moduleShortName = value;
            }
        }

        [XmlAttribute]
        public string Name { get; set; }

        [XmlAttribute]
        public string Description { get; set; }

        /// <summary>
        /// Должны быть соответствующие файлы .onepkg
        /// </summary>
        [XmlElement(typeof(NotebookInfo), ElementName = "Notebook")]
        public List<NotebookInfo> Notebooks { get; set; }

        [XmlAttribute]
        public string DictionarySectionGroupName { get; set; }

        [XmlIgnore]
        public int? StrongNumbersCount { get; set; }

        [XmlAttribute]
        [EditorBrowsable(EditorBrowsableState.Never)]
        public int XmlStrongNumbersCount
        {
            get
            {
                return StrongNumbersCount.Value;
            }
            set
            {
                StrongNumbersCount = value;
            }
        }

        [EditorBrowsable(EditorBrowsableState.Never)]
        public bool XmlStrongNumbersCountSpecified
        {
            get
            {
                return StrongNumbersCount.HasValue;
            }
        }

        /// <summary>
        /// Должны быть соответствующие файлы .one
        /// </summary>
        [XmlElement(typeof(SectionInfo), ElementName = "Section")]
        public List<SectionInfo> Sections { get; set; } 

        [XmlElement]
        public BibleTranslationDifferences BibleTranslationDifferences { get; set; }

        [XmlElement]
        public BibleStructureInfo BibleStructure { get; set; }

        public ModuleInfo()
        {
            this.BibleTranslationDifferences = new BibleTranslationDifferences();
        }

        public bool UseSingleNotebook()
        {
            return Notebooks.Exists(n => n.Type == ContainerType.Single);
        }

        public NotebookInfo GetNotebook(ContainerType notebookType)
        {
            return Notebooks.First(n => n.Type == notebookType);
        }

        /// <summary>
        /// возвращает книгу Библии с учётом всех сокращений 
        /// </summary>
        /// <param name="s"></param>
        /// <param name="endsWithDot">Методу передаётся уже стримленная строка. Потому отдельно передаётся: заканчивалось ли название книги на точку. Если имя книги было полное (а не сокращённое) и оно окончивалось на точку, то не считаем это верной записью</param>
        /// <returns></returns>
        public BibleBookInfo GetBibleBook(string s, bool endsWithDot, out string moduleName)
        {
            s = s.ToLowerInvariant();

            var result = GetBibleBookByExactMatch(s, endsWithDot, out moduleName);

            if (result == null && s.Length > 0 && StringUtils.IsDigit(s[0]))
            {
                s = s.Replace(" ", string.Empty); // чтоб находил "1 John", когда в списке сокращений только "1John"
                result = GetBibleBookByExactMatch(s, endsWithDot, out moduleName);
            }

            return result;
        }

        private BibleBookInfo GetBibleBookByExactMatch(string s, bool endsWithDot, out string moduleName)
        {
            moduleName = null;

            foreach (var book in BibleStructure.BibleBooks)
            {
                if (book.Name.ToLowerInvariant() == s || book.SectionName.ToLowerInvariant() == s)
                {
                    if (endsWithDot)
                        return null;

                    return book;
                }

                var abbreviation = book.Abbreviations.FirstOrDefault(abbr => abbr.Value == s
                                                        && (!endsWithDot || !abbr.IsFullBookName));
                if (abbreviation != null)
                {
                    moduleName = abbreviation.ModuleName;
                    return book;
                }
            }

            return null;
        }

        /// <summary>
        /// из-за проблем совместимости версий модуля
        /// </summary>
        public void CorrectModuleAfterDeserialization()
        {
            if (Version.CompareTo("2.0") < 0)
            {
                var bibleNotebook = Notebooks.FirstOrDefault(n => n.Type == ContainerType.Bible);
                if (bibleNotebook != null)
                {
                    bibleNotebook.SectionGroups = new List<SectionGroupInfo>() 
                    {        
                        new SectionGroupInfo() 
                        { 
                            Name = BibleStructure.OldTestamentName, 
                            CheckSectionsCount = true, 
                            SectionsCount = BibleStructure.OldTestamentBooksCount, 
                            Type = ContainerType.OldTestament
                        },
                        new SectionGroupInfo() 
                        { 
                            Name = BibleStructure.NewTestamentName, 
                            CheckSectionsCount = true, 
                            SectionsCount = BibleStructure.NewTestamentBooksCount, 
                            Type = ContainerType.NewTestament 
                        }
                    };
                }

                var commentsNotebooks = Notebooks.Where(n => n.Type == ContainerType.BibleComments || n.Type == ContainerType.BibleNotesPages);
                foreach (var commentsNotebook in commentsNotebooks)
                {
                    commentsNotebook.SectionGroups = new List<SectionGroupInfo>()
                    {
                        new SectionGroupInfo() 
                        { 
                            Name = BibleStructure.OldTestamentName, 
                            CheckSectionsCount = true, 
                            SectionsCountMax = 3, 
                            Type = ContainerType.OldTestament 
                        },
                        new SectionGroupInfo() 
                        { 
                            Name = BibleStructure.NewTestamentName, 
                            CheckSectionsCount = true, 
                            SectionsCountMax = 3, 
                            Type = ContainerType.NewTestament 
                        }
                    };
                }
            }
        }               
    }

    [Serializable]
    public class NotebookInfo : SectionGroupInfo
    {

    }

    [Serializable]
    public class SectionGroupInfo
    {
        [XmlAttribute]
        public ContainerType Type { get; set; }

        [XmlAttribute]
        public string Name { get; set; }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool SkipCheck { get; set; }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool CheckSectionGroupsCount { get; set; }

        [XmlAttribute]
        [DefaultValue(0)]
        public int SectionGroupsCount { get; set; }

        [XmlAttribute]
        [DefaultValue(0)]
        public int SectionGroupsCountMin { get; set; }

        [XmlAttribute]
        [DefaultValue(0)]
        public int SectionGroupsCountMax { get; set; }

        [XmlElement(typeof(SectionGroupInfo), ElementName = "SectionGroup")]
        public List<SectionGroupInfo> SectionGroups { get; set; }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool CheckSectionsCount { get; set; }

        [XmlAttribute]
        [DefaultValue(0)]
        public int SectionsCount { get; set; }

        [XmlAttribute]
        [DefaultValue(0)]
        public int SectionsCountMin { get; set; }

        [XmlAttribute]
        [DefaultValue(0)]
        public int SectionsCountMax { get; set; }

        [XmlElement(typeof(SectionInfo), ElementName = "Section")]
        public List<SectionInfo> Sections { get; set; }
    }

    [Serializable]
    public class SectionInfo
    {
        [XmlAttribute]
        public string Name { get; set; }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool SkipCheck { get; set; }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool CheckPagesCount { get; set; }

        [XmlAttribute]
        [DefaultValue(0)]
        public int PagesCount { get; set; }

        [XmlAttribute]
        [DefaultValue(0)]
        public int PagesCountMin { get; set; }

        [XmlAttribute]
        [DefaultValue(0)]
        public int PagesCountMax { get; set; }
    }

    [Serializable]
    public class BibleStructureInfo
    {
        [XmlAttribute]
        [DefaultValue("")]
        //[Obsolete()]  если пометить - то перестанут значения считываться из XML файла
        public string OldTestamentName { get; set; }

        [XmlAttribute]
        [DefaultValue(0)]
        //[Obsolete()]
        public int OldTestamentBooksCount { get; set; }

        [XmlAttribute]
        [DefaultValue("")]
        //[Obsolete()]
        public string NewTestamentName { get; set; }

        [XmlAttribute]
        [DefaultValue(0)]
        //[Obsolete()]
        public int NewTestamentBooksCount { get; set; }

        [XmlAttribute]
        public string Alphabet { get; set; }  // символы, встречающиеся в названии книг Библии                    

        [XmlAttribute]
        public string ChapterSectionNameTemplate { get; set; }

        [XmlElement(typeof(BibleBookInfo), ElementName = "BibleBook")]
        public List<BibleBookInfo> BibleBooks { get; set; }

        public BibleStructureInfo()
        {
            this.BibleBooks = new List<BibleBookInfo>();
        }
    }

    [Serializable]
    public class BibleBookInfo
    {
        [XmlAttribute]
        public int Index { get; set; }

        [XmlAttribute]
        public string Name { get; set; }

        [XmlAttribute]
        public string SectionName { get; set; }

        [XmlElement(typeof(Abbreviation), ElementName = "Abbreviation")]
        public List<Abbreviation> Abbreviations { get; set; }
    }

    [Serializable]
    public class Abbreviation
    {
        [XmlAttribute]
        [DefaultValue(false)]
        public bool IsFullBookName { get; set; }

        [XmlText]
        public string Value { get; set; }

        [XmlAttribute]
        [DefaultValue("")]
        public string ModuleName { get; set; }

        public Abbreviation()
        {
        }

        public Abbreviation(string value)
        {
            this.Value = value;
        }

        public static implicit operator Abbreviation(string value)
        {
            return new Abbreviation(value);
        }
    }

    [Serializable]
    [XmlRoot(ElementName = "IStudyBibleTools_Dictionary")]
    public class ModuleDictionaryInfo
    {
        [XmlElement]
        public TermSet TermSet { get; set; }        
    }

    [Serializable]
    public class TermSet 
    {
        [XmlElement(typeof(string), ElementName = "Term")]
        public List<string> Terms { get; set; }
    }

    [Serializable]
    [XmlRoot(ElementName = "IStudyBibleTools_Bible")]
    public class ModuleBibleInfo
    {
        [XmlElement]
        public BibleContent Content { get; set; }

        public ModuleBibleInfo()
        {
            this.Content = new BibleContent();
        }
    }

    [Serializable]
    public class BibleTranslationDifferences
    {
        /// <summary>
        /// Первые буквы алфавита для разбиения стихов на части
        /// </summary>
        [XmlAttribute]
        public string PartVersesAlphabet { get; set; }

        [XmlElement(typeof(BibleBookDifferences), ElementName = "BookDifferences")]
        public List<BibleBookDifferences> BookDifferences { get; set; }

        public BibleTranslationDifferences()
        {
            this.BookDifferences = new List<BibleBookDifferences>();
        }
    }

    [Serializable]
    public class BibleBookDifferences
    {
        [XmlAttribute]
        public int BookIndex { get; set; }

        [XmlElement(typeof(BibleBookDifference), ElementName = "Difference")]
        public List<BibleBookDifference> Differences { get; set; }

        public BibleBookDifferences()
        {
            this.Differences = new List<BibleBookDifference>();
        }

        public BibleBookDifferences(int bookIndex, params BibleBookDifference[] bibleBookDifferences)
            : this()
        {
            this.BookIndex = bookIndex;
            this.Differences.AddRange(bibleBookDifferences);
        }
    }

    [Serializable]
    public class BibleBookDifference
    {
        /// <summary>
        /// Выравнивание стихов, если, например, на два стиха приходится один параллельный
        /// </summary>
        public enum CorrespondenceVerseType
        {
            All = 0,
            First = 1,
            Last = 2
        }

        [XmlAttribute]
        public string BaseVerses { get; set; }

        [XmlAttribute]
        public string ParallelVerses { get; set; }

        [XmlAttribute]
        [DefaultValue(false)]
        public bool SkipCheck { get; set; }

        /// <summary>
        /// Выравнивание стихов - при несоответствии, 
        /// </summary>
        [XmlAttribute]
        [DefaultValue((int)CorrespondenceVerseType.All)]
        public CorrespondenceVerseType CorrespondenceType { get; set; }

        /// <summary>
        /// Количество стихов, соответствующих версии KJV. По умолчанию - все стихи соответствуют KJV (если Strict = true), либо только один стих (если Strict = false)
        /// </summary>
        [XmlAttribute]
        public string ValueVersesCount { get; set; }

        public BibleBookDifference()
        {
        }

        //public BibleBookDifference(string baseVerses, string parallelVerses)
        //    : this()
        //{
        //    this.BaseVerses = baseVerses;
        //    this.ParallelVerses = parallelVerses;
        //}
    }

    [Serializable]
    public class BibleContent
    {
        [XmlAttribute]
        public string Locale { get; set; }

        [XmlElement(typeof(BibleBookContent), ElementName = "Book")]
        public List<BibleBookContent> Books { get; set; }

        public BibleContent()
        {
            this.Books = new List<BibleBookContent>();
        }
    }

    [Serializable]
    public class BibleBookContent
    {
        [XmlAttribute]
        public int Index { get; set; }

        [XmlElement(typeof(BibleChapterContent), ElementName = "Chapter")]
        public List<BibleChapterContent> Chapters { get; set; }

        public BibleBookContent()
        {
            this.Chapters = new List<BibleChapterContent>();
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

    [Serializable]
    public class BibleChapterContent
    {
        [XmlAttribute]
        public int Index { get; set; }

        [XmlElement(typeof(BibleVerseContent), ElementName = "Verse")]
        public List<BibleVerseContent> Verses { get; set; }

        private Dictionary<int, BibleVerseContent> _versesDictionary;
        public BibleVerseContent GetVerse(int verseNumber)
        {
            if (_versesDictionary == null)            
                LoadVersesDictionary();

            if (_versesDictionary.ContainsKey(verseNumber))
                return _versesDictionary[verseNumber];
            else if (Verses.Any(v => v.IsMultiVerse && (v.Index <= verseNumber && verseNumber <= v.TopIndex)))
                return BibleVerseContent.Empty;

            return null;
        }

        private void LoadVersesDictionary()
        {
            _versesDictionary = new Dictionary<int,BibleVerseContent>();
            foreach (var verse in Verses)
                _versesDictionary.Add(verse.Index, verse);            
        }

        public BibleChapterContent()
        {
            this.Verses = new List<BibleVerseContent>();
        }
    }

    [Serializable]
    public class BibleVerseContent
    {
        [XmlAttribute]
        public int Index { get; set; }

        [XmlIgnore]
        public int? TopIndex { get; set; }

        [XmlIgnore]
        public bool IsMultiVerse { get { return TopIndex.HasValue; } }

        [XmlIgnore]
        public VerseNumber VerseNumber
        {
            get
            {
                return new VerseNumber(Index, TopIndex);
            }
        }

        [XmlText]
        public string Value { get; set; }

        public static BibleVerseContent Empty
        {
            get
            {
                return new BibleVerseContent();
            }
        }

        [XmlIgnore]
        public bool IsEmpty
        {
            get
            {
                return Index == default(int);
            }
        }

        [XmlAttribute]
        [EditorBrowsable(EditorBrowsableState.Never)]
        public int XmlTopIndex
        {
            get
            {
                return TopIndex.Value;
            }
            set
            {
                TopIndex = value;
            }
        }

        [EditorBrowsable(EditorBrowsableState.Never)]
        public bool XmlTopIndexSpecified
        {
            get
            {
                return TopIndex.HasValue;
            }
        }
    }
}