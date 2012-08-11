using System;
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
        public BibleTranslationDifferences BibleTranslationDifferences { get; set; }

        [XmlElement]
        public BibleStructureInfo BibleStructure { get; set; }               

        public ModuleInfo()
        {
            this.BibleTranslationDifferences = new BibleTranslationDifferences();
        }

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
        public enum VerseAlign
        {
            None = 0,
            Top = 1,
            Bottom = 2
        }

        [XmlAttribute]
        public string BaseVerses { get; set; }

        [XmlAttribute]
        public string ParallelVerses { get; set; }

        /// <summary>
        /// Выравнивание стихов - при несоответствии, 
        /// </summary>
        [XmlAttribute]
        [DefaultValue((int)VerseAlign.None)]
        public VerseAlign Align { get; set; }

        /// <summary>
        /// Строгая обработка стихов - отслеживается строгое соответствие
        /// </summary>
        [XmlAttribute]
        [DefaultValue(true)]  
        public bool Strict { get; set; }

        /// <summary>
        /// Количество стихов, соответствующие частям из KJV. Например при "1:1 -> 1:1-3", 
        /// и если 1:1 делится только на две части с помощью "|", то надр, чтобы ValueVerseCount=2 и, например, Align = Bottom. 
        /// Тогда 2 и 3 стих будут соответствовать 1:1 из KJV, а 1 стих - "особенный", который есть только в данном переводе
        /// По умолчанию ValueVerseCount=null, то есть все стихи соответствуют частям/стихам из KJV
        /// Данный параметр полезен для апокрифов
        /// </summary>
        [XmlAttribute]
        public string ValueVerseCount { get; set; }  

        public BibleBookDifference()
        {
            this.Strict = true;
        }

        public BibleBookDifference(string baseVerses, string parallelVerses)
            : this()
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

        public string GetVerseContent(SimpleVersePointer verse)
        {
            if (this.Chapters.Count < verse.Chapter)
                throw new ArgumentException(string.Format("There is no verse '{0}'", verse));
            
            var chapter = this.Chapters[verse.Chapter - 1];

            if (chapter.Verses.Count < verse.Verse)
                throw new ArgumentException(string.Format("There is no verse '{0}'", verse));

            string verseContent = chapter.Verses[verse.Verse - 1].Value;

            string result = null;

            if (verse.PartIndex.HasValue)
            {
                var versesParts = verseContent.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                if (versesParts.Length > verse.PartIndex.Value)
                    result = versesParts[verse.PartIndex.Value];                
            }
            else
                result = verseContent;

            if (!string.IsNullOrEmpty(result))
                result = result.Replace("|", string.Empty);

            return result;
        }

        public string GetVersesContent(List<SimpleVersePointer> verses)
        {
            StringBuilder versesContent = new StringBuilder();

            foreach (var verse in verses)
            {
                versesContent.Append(GetVerseContent(verse));
            }

            return versesContent.ToString();
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