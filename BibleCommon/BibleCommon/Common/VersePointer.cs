using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections.Specialized;
using BibleCommon.Consts;
using BibleCommon.Helpers;
using BibleCommon.Services;
using Microsoft.Office.Interop.OneNote;
using BibleCommon.Contracts;

namespace BibleCommon.Common
{
    [Serializable]
    public struct VerseNumber
    {
        public int Verse;
        public int? TopVerse;
        public bool IsMultiVerse { get { return TopVerse.HasValue; } }

        public bool IsVerseBelongs(int verse)
        {
            if (!IsMultiVerse)
                return Verse == verse;
            else
                return Verse <= verse && verse <= TopVerse.Value; 
        }

        public VerseNumber(int verse)            
        {
            Verse = verse;
            TopVerse = null;
        }

        public VerseNumber(int verse, int? topVerse)
        {
            Verse = verse;
            if (topVerse.GetValueOrDefault(-1) > Verse)
                TopVerse = topVerse;
            else
                TopVerse = null;
        }

        public List<int> GetAllVerses()
        {
            var result = new List<int>();

            result.Add(Verse);

            if (IsMultiVerse)
            {
                for (int i = Verse + 1; i <= TopVerse; i++)
                    result.Add(i);
            }

            return result;
        }

        public static VerseNumber Parse(string s)
        {
            s = s.Trim();
            var parts = s.Split(new char[] { '-' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 1)
                return new VerseNumber(int.Parse(s));
            else if (parts.Length == 2)
                return new VerseNumber(int.Parse(parts[0]), int.Parse(parts[1]));
            else
                throw new NotSupportedException(s);
        }

        public static VerseNumber? GetFromVerseText(string verseText)
        {
            string temp;
            return GetFromVerseText(verseText, out temp);
        }

        public static VerseNumber? GetFromVerseText(string verseText, out string verseTextWithoutNumber)
        {
            verseText = verseText.Trim();
            verseTextWithoutNumber = null;
            int textBreakIndex, htmlBreakIndex;
            var verseIndex = StringUtils.GetNextString(verseText, -1, new SearchMissInfo(0, SearchMissInfo.MissMode.CancelOnMissFound),
                                                        out textBreakIndex, out htmlBreakIndex, StringSearchIgnorance.None, StringSearchMode.SearchNumber);
            if (!string.IsNullOrEmpty(verseIndex))
            {
                int? topVerse = null;
                if (StringUtils.GetChar(verseText, htmlBreakIndex) == '-')
                {
                    var topVerseString = StringUtils.GetNextString(verseText, htmlBreakIndex, new SearchMissInfo(0, SearchMissInfo.MissMode.CancelOnMissFound),
                                                        out textBreakIndex, out htmlBreakIndex, StringSearchIgnorance.None, StringSearchMode.SearchNumber);
                    if (!string.IsNullOrEmpty(topVerseString))
                        topVerse = int.Parse(topVerseString);
                }                

                verseTextWithoutNumber = verseText.Substring(htmlBreakIndex + 1);

                return new VerseNumber(int.Parse(verseIndex), topVerse);
            }

            return null;
        }

        public override string ToString()
        {
            if (IsMultiVerse)
                return string.Format("{0}-{1}", Verse, TopVerse);
            else
                return Verse.ToString();
        }

        public override int GetHashCode()
        {
            var result = Verse.GetHashCode();
            if (TopVerse.HasValue)
                result = result ^ TopVerse.Value.GetHashCode();

            return result;
        }

        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;

            if (!(obj is VerseNumber))
                return false;

            var anotherObj = (VerseNumber)obj;

            return this.Verse == anotherObj.Verse 
                && this.TopVerse == anotherObj.TopVerse;
        }

        public static bool operator ==(VerseNumber vn1, VerseNumber vn2)
        {
            if (((object)vn1) == null && ((object)vn2) == null)
                return true;

            if (((object)vn1) == null)
                return false;

            if (((object)vn2) == null)
                return false;

            return vn1.Equals(vn2);
        }

        public static bool operator !=(VerseNumber vn1, VerseNumber vn2)
        {
            return !(vn1 == vn2);
        }        
    }

    [Serializable]
    public class SimpleVersePointer: ICloneable
    {
        public int BookIndex { get; set; }
        public int Chapter { get; set; }
        public VerseNumber? VerseNumber { get; set; }        
        public int? PartIndex { get; set; }        
        public bool IsEmpty { get; set; }
        public bool IsApocrypha { get; set; }
        public bool SkipCheck { get; set; }

        public bool IsMultiVerse
        {
            get
            {
                return VerseNumber.HasValue && VerseNumber.Value.IsMultiVerse;
            }
        }        

        public int Verse
        {
            get
            {
                if (VerseNumber.HasValue)
                    return VerseNumber.Value.Verse;

                return default(int);
            }
        }

        public int? TopVerse
        {
            get
            {
                if (VerseNumber.HasValue)
                    return VerseNumber.Value.TopVerse;

                return null;
            }
        }

        public bool IsChapter
        {
            get
            {
                return this.VerseNumber == null;
            }
        }

        public SimpleVersePointer(string s)
        {
            var parts = s.Split(new char[] { ' ', ':' });

        }

        public SimpleVersePointer()
        {
        }

        public SimpleVersePointer(SimpleVersePointer verse)
            : this(verse.BookIndex, verse.Chapter, new VerseNumber(verse.Verse, verse.TopVerse))
        {

        }

        public SimpleVersePointer(int bookIndex, int chapter)
            : this(bookIndex, chapter, null)
        { }

        public SimpleVersePointer(int bookIndex, int chapter, VerseNumber? verse)
        {
            this.BookIndex = bookIndex;
            this.Chapter = chapter;
            this.VerseNumber = verse;
        }

        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;

            if (!(obj is SimpleVersePointer))
                return false;

            SimpleVersePointer other = (SimpleVersePointer)obj;            
            return this.BookIndex == other.BookIndex
                && this.Chapter == other.Chapter
                && this.VerseNumber == other.VerseNumber;
        }

        public override int GetHashCode()
        {
            return this.BookIndex.GetHashCode() ^ this.Chapter.GetHashCode() ^ this.VerseNumber.GetHashCode();
        }

        public override string ToString()
        {
            var result = string.Format("{0} {1}:{2}", BookIndex, Chapter, VerseNumber);

            if (PartIndex.HasValue)
                result += string.Format("({0})", PartIndex);

            if (IsEmpty)
                result += "(empty)";

            if (IsApocrypha)
                result += "(A)";

            return result;
        }        

        public virtual object Clone()
        {
            var result = new SimpleVersePointer(this);
            CopyProperties(result);

            return result;
        }

        protected void CopyProperties(SimpleVersePointer verse)
        {
            verse.IsApocrypha = this.IsApocrypha;
            verse.IsEmpty = this.IsEmpty;
            verse.PartIndex = this.PartIndex;            
            verse.SkipCheck = this.SkipCheck;
        }

        public SimpleVersePointer GetChapterPointer()
        {
            return new SimpleVersePointer(this.BookIndex, this.Chapter);
        }

        public List<SimpleVersePointer> GetAllVerses()
        {
            var result = new List<SimpleVersePointer>();

            if (this.VerseNumber.HasValue)
            {
                result.AddRange(this.VerseNumber.Value.GetAllVerses().ConvertAll(v =>
                {
                    var verse = (SimpleVersePointer)this.Clone();
                    verse.VerseNumber = new VerseNumber(v);
                    return verse;
                }));
            }
            else
                result.Add(this);

            return result;
        }

        public VersePointer ToVersePointer(ModuleInfo moduleInfo)
        {
            var bookInfo = moduleInfo.BibleStructure.BibleBooks.FirstOrDefault(book => book.Index == this.BookIndex);
            if (bookInfo == null)
                throw new ArgumentException(string.Format("Book with index {0} was not found in module {1}", this.BookIndex, moduleInfo.ShortName));

            return new VersePointer(bookInfo.Name, this.Chapter, this.Verse, this.TopVerse);
        }        
    }

    public class SimpleVerse : SimpleVersePointer
    {

        /// <summary>
        /// Строка, соответствующая номерам/номеру стиха. Может быть: "5", "5-6", "6:5-6"
        /// </summary>
        public string VerseNumberString { get; set; }

        /// <summary>
        /// Ссылка, ведующая с номера стиха
        /// </summary>
        public string VerseLink { get; set; }

        /// <summary>
        /// Текст стиха без номера
        /// </summary>
        public string VerseContent { get; set; }

        public string GetVerseFullString()
        {
            if (this.IsApocrypha || this.IsEmpty)
                return string.Empty;

            string verseNumber = string.IsNullOrEmpty(VerseLink) 
                                    ? this.VerseNumberString
                                    : string.Format("<a href='{0}'>{1}</a>", VerseLink, VerseNumberString);

            return string.Format("{0}{1}{2}",
                            verseNumber,
                            string.IsNullOrEmpty(VerseContent) ? string.Empty : "<span> </span>",
                            VerseContent);
        }

        public SimpleVerse(SimpleVersePointer versePointer, string verseContent)
            : this(versePointer, null, verseContent)
        { }
     
        /// <summary>
        /// string.IsNullOrEmpty(verseContent) - это не то же самое, что this.IsEmpty
        /// </summary>
        /// <param name="versePointer"></param>
        /// <param name="verseContent"></param>
        public SimpleVerse(SimpleVersePointer versePointer, string verseNumber, string verseContent)
            : base(versePointer.BookIndex, versePointer.Chapter, versePointer.VerseNumber)
        {
            this.VerseContent = verseContent;

            if (!string.IsNullOrEmpty(verseNumber))
                this.VerseNumberString = verseNumber;
            else
                this.VerseNumberString = versePointer.Verse.ToString();            
        }

        public override object Clone()
        {
            var result = new SimpleVerse(this, this.VerseNumberString, this.VerseContent);
            CopyProperties(result);

            return result;            
        }
    }

    [Serializable]
    public class VersePointer
    {      
        public BibleBookInfo Book { get; set; }
        public int? Chapter { get; set; }
        public int? Verse { get; set; }
        public string OriginalVerseName { get; set; }   // первоначально переданная строка в конструктор
        public string OriginalBookName { get; set; }

        public VersePointer ParentVersePointer { get; set; } // родительская ссылка. Например если мы имеем дело со стихом диапазона, то здесь хранится стих, являющийся диапазоном

        public VerseNumber VerseNumber
        {
            get
            {
                return new VerseNumber(this.Verse.GetValueOrDefault(), this.TopVerse);
            }
        }

        public override string ToString()
        {
            if (!string.IsNullOrEmpty(this.OriginalVerseName))
                return this.OriginalVerseName;

            return base.ToString();
        }

        public static VersePointer GetChapterVersePointer(string bookName, int chapter)
        {
            return GetChapterVersePointer(string.Format("{0} {1}", bookName, chapter));
        }

        public static VersePointer GetChapterVersePointer(string chapterName)
        {
            return new VersePointer(string.Format("{0}:0", chapterName));
        }

        public string FriendlyVerseName
        {
            get
            {
                return string.Format("{0}:{1}", ChapterName, Verse);
            }
        }    

        public string ChapterName
        {
            get
            {                
                return string.Format("{0} {1}", Book != null ? Book.Name : string.Empty, Chapter);
            }
        }

        public VersePointer(VersePointer chapterPointer, int verse)
            : this(string.Format("{0} {1}:{2}", chapterPointer.OriginalBookName, chapterPointer.Chapter, verse))
        {

        }

        public VersePointer(string bookName, int chapter)
            : this(bookName, chapter, 0)
        {

        }

        public VersePointer(string bookName, int chapter, int verse)
            : this(bookName, chapter, verse, null)
        {

        }

        public VersePointer(string bookName, int chapter, int verse, int? topVerse)
            : this(string.Format("{0} {1}:{2}{3}{4}", bookName, chapter, verse, topVerse.HasValue ? "-" : string.Empty, topVerse))
        {

        }

        public VersePointer(string s)
        {
            if (!string.IsNullOrEmpty(s))
            {
                this.OriginalVerseName = s;

                s = s.ToLowerInvariant();

                s = TrimLocation(s);

                int? verse = StringUtils.GetStringLastNumber(s);

                if (verse.HasValue)
                {
                    Verse = verse.Value;

                    int i = s.LastIndexOf(Verse.ToString());

                    s = s.Substring(0, i);

                    int chapterIndex;
                    int? chapter = StringUtils.GetStringLastNumber(s, out chapterIndex);

                    if (chapter.HasValue && chapterIndex > 1)  // чтобы уберечь от строк типа "1Кор" - то есть считаем, что глава идёт хотя бы с третьего символа
                    {
                        Chapter = chapter.Value;

                        i = s.LastIndexOf(Chapter.ToString());

                        s = s.Substring(0, i);
                    }
                    else
                    {
                        Chapter = Verse;
                        Verse = 0;
                        _topChapter = _topVerse;
                        _topVerse = null;
                    }

                    if (
                        ((TopChapter == null || TopChapter.GetValueOrDefault(0) == Chapter.GetValueOrDefault(0))
                                        && TopVerse.GetValueOrDefault(0) <= Verse.GetValueOrDefault(0))
                        || 
                        (TopChapter != null 
                                        && TopChapter.GetValueOrDefault(0) < Chapter.GetValueOrDefault(0)))
                    {
                        _topVerse = null;
                        _topChapter = null;
                        _isMultiVerse = false;
                    }                    
                }

                bool endsWithDot;
                string moduleName;
                OriginalBookName = TrimBookName(s, out endsWithDot);
                Book = GetBibleBook(OriginalBookName, endsWithDot, out moduleName);

                if (!string.IsNullOrEmpty(moduleName))   // значит ссылка дана для модуля, отличного от установленного                
                    ConvertToBaseVerse(moduleName);                
            }
        }       

        public SimpleVersePointer ToSimpleVersePointer()
        {
            return new SimpleVersePointer(Book.Index, Chapter.GetValueOrDefault(0), new VerseNumber(Verse.GetValueOrDefault(0), this.TopVerse));
        }

        private void ConvertToBaseVerse(string moduleName)
        {
            if (IsValid)
            {
                var parallelVersePointer = BibleParallelTranslationConnectorManager.GetParallelVersePointer(
                                                ToSimpleVersePointer(), moduleName, SettingsManager.Instance.ModuleName);

                this.OriginalBookName = this.Book.Name;
                this.Chapter = parallelVersePointer.Chapter;
                this.Verse = parallelVersePointer.Verse;

                if (IsMultiVerse)
                {
                    parallelVersePointer = BibleParallelTranslationConnectorManager.GetParallelVersePointer(
                                                new SimpleVersePointer(
                                                                    this.Book.Index, 
                                                                    this.TopChapter.GetValueOrDefault(this.Chapter.Value), 
                                                                    new VerseNumber(this.TopVerse.GetValueOrDefault(this.Verse.Value))), 
                                                moduleName, SettingsManager.Instance.ModuleName);

                    if (TopChapter.HasValue)
                        _topChapter = parallelVersePointer.Chapter;
                    if (TopVerse.HasValue)
                        _topVerse = parallelVersePointer.Verse;
                }
            }
        }

        public bool IsChapter
        {
            get
            {
                return IsVerseChapter(this.Verse);
            }
        }

        public static bool IsVerseChapter(int? verseNumber)
        {
            return verseNumber.GetValueOrDefault(0) == 0;
        }


        private bool _isMultiVerse = false;
        /// <summary>
        /// охватывает несколько стихов, например Откр 5:6-9
        /// </summary>
        public bool IsMultiVerse  
        {
            get
            {
                return _isMultiVerse;
            }
        }

        private int? _topVerse = null;
        /// <summary>
        /// Верхняя граница ряда стихов. Например стих 9 в Откр 5:6-9
        /// </summary>        
        public int? TopVerse
        {
            get
            {
                return _topVerse;
            }
        }


        private int? _topChapter = null;
        public int? TopChapter
        {
            get
            {
                return _topChapter;
            }
        }

        /// <summary>
        /// Принадлежит ли указанный стих текущему диапазону стихов
        /// </summary>
        /// <param name="verse"></param>
        /// <returns></returns>
        public bool IsInVerseRange(int verse)
        {
            if (IsMultiVerse)
                return verse >= this.Verse && verse <= this.TopVerse;

            return false;
        }


        private string TrimLocation(string s)
        {            
            int indexOfDash = s.LastIndexOf("-");

            if (indexOfDash != -1 && indexOfDash > 3) // чтоб отсечь такие варианты, как "2-Тим 2:3"
            {
                int tempTopVerse;
                string topStructure = s.Substring(indexOfDash + 1);
                if (int.TryParse(topStructure, out tempTopVerse))
                {
                    _topVerse = tempTopVerse;
                    _isMultiVerse = true;
                }
                else
                {
                    int indexOfDelimiter = topStructure.IndexOf(":");
                    if (indexOfDelimiter != -1)  // значит скорее всего имеем дело с Ин 1:5-2:4
                    {
                        int tempTopChapter;
                        if (int.TryParse(topStructure.Substring(0, indexOfDelimiter), out tempTopChapter)
                            && int.TryParse(topStructure.Substring(indexOfDelimiter + 1), out tempTopVerse))
                        {
                            _topChapter = tempTopChapter;
                            _topVerse = tempTopVerse;
                            _isMultiVerse = true;
                        }
                    }
                }

                s = s.Substring(0, indexOfDash);
            }

            int iComma = s.LastIndexOf(',');
            if (iComma != -1)
            {
                int iDel = s.IndexOfAny(new char[] { '.', ',', ':' });
                if (iDel != -1)
                {
                    if (iComma > iDel) // ситуация типа "2Тим 1:13,14"
                        s = s.Substring(0, iComma);
                }
            }

            return s;
        }

        private static string TrimBookName(string bookName, out bool endsWithDot)
        {
            string result = bookName.Trim();
            endsWithDot = result.EndsWith(".");

            return result.Trim('.').Replace("  ", " ");              
        }

        public bool IsValid
        {
            get
            {
                return Book != null                    
                    && Chapter.HasValue
                    && Verse.HasValue;
            }
        }

        private static BibleBookInfo GetBibleBook(string s, bool endsWithDot, out string moduleName)
        {   
            return SettingsManager.Instance.CurrentModule.GetBibleBook(s, endsWithDot, out moduleName);
        }

        public VersePointer GetChapterPointer()
        {
            return new VersePointer(this.OriginalBookName, this.Chapter.Value, 0);
        }
        

        /// <summary>
        /// возвращает список всех вложенных стихов (за исключением первого) если Multiverse. 
        /// </summary>
        /// <returns></returns>
        public List<VersePointer> GetAllIncludedVersesExceptFirst(Application oneNoteApp, GetAllIncludedVersesExceptFirstArgs args)
        {
            List<VersePointer> result = new List<VersePointer>();
            if ((IsValid || args.Force) && IsMultiVerse)
            {
                if (TopChapter != null && TopVerse != null && !args.SearchOnlyForFirstChapter)
                {
                    for (int chapterIndex = Chapter.Value; chapterIndex <= TopChapter; chapterIndex++)
                    {
                        int topVerse, startVerseIndex;
                        if (chapterIndex == TopChapter)
                            topVerse = TopVerse.Value;
                        else
                            topVerse = HierarchySearchManager.GetChapterVersesCount(
                                            oneNoteApp, args.BibleNotebookId, 
                                            VersePointer.GetChapterVersePointer(this.OriginalBookName, chapterIndex))
                                            .GetValueOrDefault(0);

                        if (chapterIndex == Chapter)
                            startVerseIndex = Verse.Value + 1;
                        else
                            startVerseIndex = 1;

                        for (int verseIndex = startVerseIndex; verseIndex <= topVerse; verseIndex++)
                        {
                            VersePointer vp = new VersePointer(this.OriginalBookName, chapterIndex, verseIndex);
                            vp.ParentVersePointer = this;
                            result.Add(vp);
                        }
                    }
                }
                else if (TopChapter != null && IsChapter)
                {
                    for (int chapterIndex = Chapter.Value + 1; chapterIndex <= TopChapter; chapterIndex++)
                    {
                        VersePointer vp = VersePointer.GetChapterVersePointer(this.OriginalBookName, chapterIndex);
                        vp.ParentVersePointer = this;
                        result.Add(vp);
                    }
                }
                else
                {
                    for (int verseIndex = Verse.GetValueOrDefault(0) + 1; verseIndex <= TopVerse; verseIndex++)
                    {
                        VersePointer vp = new VersePointer(this.OriginalBookName, this.Chapter.Value, verseIndex);
                        vp.ParentVersePointer = this;
                        result.Add(vp);
                    }
                }
            }

            return result;
        }

        public override bool Equals(object obj)
        {
            VersePointer otherVp = (VersePointer)obj;
            return this.Book != null && otherVp.Book != null
                && this.Book.Name == otherVp.Book.Name
                && this.Chapter == otherVp.Chapter
                && this.Verse == otherVp.Verse;                
        }

        public override int GetHashCode()
        {
            var result =  this.Chapter.GetValueOrDefault(0).GetHashCode() ^ this.Verse.GetValueOrDefault(0).GetHashCode();
            if (this.Book != null)
                result = result ^ this.Book.Name.GetHashCode();

            return result;
        }

        public static bool operator ==(VersePointer vp1, VersePointer vp2)
        {
            if (((object)vp1) == null && ((object)vp2) == null)
                return true;

            if (((object)vp1) == null)
                return false;

            if (((object)vp2) == null)
                return false;

            return vp1.Equals(vp2);
        }

        public static bool operator !=(VersePointer vp1, VersePointer vp2)
        {
            return !(vp1 == vp2);
        }        
    }
}
