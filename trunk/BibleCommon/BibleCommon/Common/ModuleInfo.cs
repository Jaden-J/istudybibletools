﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;
using BibleCommon.Helpers;
using System.ComponentModel;

namespace BibleCommon.Common
{
    public enum NotebookType
    {
        Single,
        Bible,
        BibleComments,
        BibleNotesPages,
        BibleStudy
    }

    public enum SectionGroupType
    {
        Bible,
        BibleComments,
        BibleStudy        
    }

    

    [Serializable]
    [XmlRoot(ElementName="IStudyBibleTools_Module")]
    public class ModuleInfo
    {
        [XmlAttribute]
        public string Version { get; set; }

        [XmlIgnore]
        public string ShortName { get; set; }

        [XmlAttribute]
        public string Name { get; set; }

        [XmlElement(typeof(NotebookInfo), ElementName = "Notebook")]
        public List<NotebookInfo> Notebooks { get; set; }

        [XmlElement]
        public BibleStructureInfo BibleStructure { get; set; }

        public bool UseSingleNotebook()
        {
            return Notebooks.Exists(n => n.Type == NotebookType.Single);
        }

        public NotebookInfo GetNotebook(NotebookType notebookType)
        {
            return Notebooks.First(n => n.Type == notebookType);
        }
        
        /// <summary>
        /// возвращает книгу Библии с учётом всех сокращений 
        /// </summary>
        /// <param name="s"></param>
        /// <param name="endsWithDot">Методу передаётся уже стримленная строка. Потому отдельно передаётся: заканчивалось ли название книги на точку. Если имя книги было полное (а не сокращённое) и оно окончивалось на точку, то не считаем это верной записью</param>
        /// <returns></returns>
        public BibleBookInfo GetBibleBook(string s, bool endsWithDot)
        {
            s = s.ToLowerInvariant();

            var result = GetBibleBookByExactMatch(s, endsWithDot);

            if (result == null && s.Length > 0 && StringUtils.IsDigit(s[0]))
            {

                s = s.Replace(" ", string.Empty); // чтоб находил "1 John", когда в списке сокращений только "1John"
                result = GetBibleBookByExactMatch(s, endsWithDot);
            }

            if (result != null && endsWithDot)
                if (result.Name.ToLowerInvariant() == s)
                    result = null;

            return result;
        }

        private BibleBookInfo GetBibleBookByExactMatch(string s, bool endsWithDot)
        {
            return BibleStructure.BibleBooks.FirstOrDefault(
                book => book.Abbreviations.Any(abbr => 
                           abbr.Value == s
                        && (!endsWithDot || !abbr.IsFullBookName)));
        }
    }

    [Serializable]    
    public class NotebookInfo
    {
        [XmlAttribute]
        public NotebookType Type { get; set; }

        [XmlAttribute]
        public string Name { get; set; }

        [XmlElement(typeof(SectionGroupInfo), ElementName="SectionGroup")]
        public List<SectionGroupInfo> SectionGroups { get; set; }
    }

    [Serializable]
    public class SectionGroupInfo
    {
        [XmlAttribute]
        public SectionGroupType Type { get; set; }

        [XmlAttribute]
        public string Name { get; set; }        
    }

    [Serializable]
    public class BibleStructureInfo
    {
        [XmlAttribute]
        public string OldTestamentName { get; set; }

        [XmlAttribute]
        public int OldTestamentBooksCount { get; set; }

        [XmlAttribute]
        public string NewTestamentName { get; set; }

        [XmlAttribute]
        public int NewTestamentBooksCount { get; set; }        

        [XmlAttribute]
        public string Alphabet { get; set; }  // символы, встречающиеся в названии книг Библии            

        [XmlElement(typeof(BibleBookInfo), ElementName="BibleBook")]
        public List<BibleBookInfo> BibleBooks { get; set; }
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
        [DefaultValueAttribute(false)]
        public bool IsFullBookName { get; set; }

        [XmlText]
        public string Value { get; set; }

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
    [XmlRoot(ElementName = "IStudyBibleTools_Bible")]
    public class ModuleBibleInfo
    {
        [XmlElement]
        public BibleTranslationDifferences TranslationDifferences { get; set; }

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
        [XmlAttribute]
        public string BaseVerses { get; set; }

        [XmlAttribute]
        public string ParallelVerses { get; set; }

        public BibleBookDifference()
        {
        }

        public BibleBookDifference(string baseVerses, string parallelVerses)
        {
            this.BaseVerses = baseVerses;
            this.ParallelVerses = parallelVerses;
        }
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

        [XmlElement(typeof(BibleChapterContent), ElementName="Chapter")]
        public List<BibleChapterContent> Chapters { get; set; }        

        public BibleBookContent()
        {
            this.Chapters = new List<BibleChapterContent>();
        }
    }

    [Serializable]
    public class BibleChapterContent
    {
        [XmlAttribute]
        public int Index { get; set; }

        [XmlElement(typeof(BibleVerseContent), ElementName = "Verse")]
        public List<BibleVerseContent> Verses { get; set; }

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

        [XmlText]
        public string Value { get; set; }
    }
}