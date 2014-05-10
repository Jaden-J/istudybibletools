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
    public struct VerseNumber: IComparable<VerseNumber>
    {
        public readonly static char[] Dashes = new char[] { '-', '—', '‑', '–', '−' };

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
            var parts = s.Split(Dashes, StringSplitOptions.RemoveEmptyEntries);
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

            var verseIndexResult = StringSearcher.SearchInString(verseText, -1, StringSearcher.SearchDirection.Forward, StringSearcher.StringSearchMode.SearchNumber,
                new StringSearcher.SearchMissInfo(0, StringSearcher.SearchMissInfo.MissMode.CancelOnMissFound));
            var htmlBreakIndex = verseIndexResult.HtmlBreakIndex;
                                                        
            if (!string.IsNullOrEmpty(verseIndexResult.FoundString))
            {
                int? topVerse = null;
                if (Dashes.Contains(StringUtils.GetChar(verseText, htmlBreakIndex)))
                {
                    var topVerseStringResult = StringSearcher.SearchInString(verseText, htmlBreakIndex, StringSearcher.SearchDirection.Forward,
                        StringSearcher.StringSearchMode.SearchNumber, new StringSearcher.SearchMissInfo(0, StringSearcher.SearchMissInfo.MissMode.CancelOnMissFound));
                    htmlBreakIndex = topVerseStringResult.HtmlBreakIndex;

                    if (!string.IsNullOrEmpty(topVerseStringResult.FoundString))
                        topVerse = int.Parse(topVerseStringResult.FoundString);
                }                

                if (verseText.Length > htmlBreakIndex + 1)
                    verseTextWithoutNumber = verseText.Substring(htmlBreakIndex + 1);

                return new VerseNumber(int.Parse(verseIndexResult.FoundString), topVerse);
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
            //if (TopVerse.HasValue)
            //    result = result ^ TopVerse.Value.GetHashCode();

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
                //&& this.TopVerse == anotherObj.TopVerse
                ;
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

        public static bool operator >(VerseNumber vn1, VerseNumber vn2)
        {
            return vn1.CompareTo(vn2) > 0;
        }

        public static bool operator >=(VerseNumber vn1, VerseNumber vn2)
        {
            return vn1.CompareTo(vn2) >= 0;
        }

        public static bool operator <(VerseNumber vn1, VerseNumber vn2)
        {
            return vn1.CompareTo(vn2) < 0;
        }

        public static bool operator <=(VerseNumber vn1, VerseNumber vn2)
        {
            return vn1.CompareTo(vn2) <= 0;
        }

        public int CompareTo(VerseNumber other)
        {
            return this.Verse.CompareTo(other.Verse);
        }
    }

    /// <summary>
    /// Важно: здесь не может быть указана TopChapter, потому не умеет работать со стихами типа "Быт 5:6-7:8". Только в пределах одной главы
    /// </summary>
    [Serializable]    
    public class SimpleVersePointer: ICloneable
    {
        public int BookIndex { get; set; }
        public int Chapter { get; set; }
        public VerseNumber VerseNumber { get; set; }        
        public int? PartIndex { get; set; }        
        public bool IsEmpty { get; set; }            // если true - то стих весь пустой: и текст и номер. Отображается пустая ячейка.
        public bool EmptyVerseContent { get; set; }  // если true - то стихи пустые, и это правильно. Если же false и стих пустой, то это ошибка.
        public bool IsApocrypha { get; set; }
        public bool SkipCheck { get; set; }

        /// <summary>
        /// Часть "бОльшего стиха". Например, если стих :3 а в ibs :2-4 - это один стих. Используется только в одном месте, не везде может быть правильно инициилизировано.
        /// </summary>
        public bool IsPartOfBigVerse { get; set; }

        /// <summary>
        /// У нас есть стих в ibs (Лев 12:7). Ему по смыслу соответствуют два стиха из rst (Лев 12:7-8). Но поделить стих в ibs не поулчается, потому палочка стоит в конце стиха. Но это не значит, что воьсмой стих пустой!
        /// </summary>
        public bool HasValueEvenIfEmpty { get; set; }

        public bool IsMultiVerse
        {
            get
            {
                return VerseNumber.IsMultiVerse;
            }
        }        

        public int Verse
        {
            get
            {   
                return VerseNumber.Verse;             
            }
        }

        public int? TopVerse
        {
            get
            {   
                return VerseNumber.TopVerse;
            }
        }

        public bool IsChapter
        {
            get
            {
                return this.VerseNumber.Verse == 0;
            }
        }

        public SimpleVersePointer GetChangedVerseAsOneChapteredBook()
        {
            return new SimpleVersePointer(this.BookIndex, 1, new VerseNumber(this.Chapter));            
        }    


        //public SimpleVersePointer(string s)
        //{
        //    var parts = s.Split(new char[] { ' ', ':' });

        //}

        public SimpleVersePointer()
        {
        }

        public SimpleVersePointer(SimpleVersePointer verse)
            : this(verse.BookIndex, verse.Chapter, new VerseNumber(verse.Verse, verse.TopVerse))
        {

        }

        public SimpleVersePointer(int bookIndex, int chapter)
            : this(bookIndex, chapter, new VerseNumber())
        { }

        public SimpleVersePointer(int bookIndex, int chapter, VerseNumber verse)
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
                && this.Verse == other.Verse;
        }

        public override int GetHashCode()
        {
            return this.BookIndex.GetHashCode() ^ this.Chapter.GetHashCode() ^ this.Verse.GetHashCode();
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

        public string ToFirstVerseString()
        {
            return string.Format("{0} {1}:{2}", BookIndex, Chapter, VerseNumber.Verse);
        }

        public virtual object Clone()
        {
            var result = new SimpleVersePointer(this);
            CopyPropertiesTo(result);

            return result;
        }

        protected void CopyPropertiesTo(SimpleVersePointer verse)
        {
            verse.IsApocrypha = this.IsApocrypha;
            verse.IsEmpty = this.IsEmpty;
            verse.PartIndex = this.PartIndex;            
            verse.SkipCheck = this.SkipCheck;
            verse.EmptyVerseContent = this.EmptyVerseContent;
            verse.IsPartOfBigVerse = this.IsPartOfBigVerse;
            verse.HasValueEvenIfEmpty = this.HasValueEvenIfEmpty;
        }

        public SimpleVersePointer GetChapterPointer()
        {
            return new SimpleVersePointer(this.BookIndex, this.Chapter);
        }

        public List<SimpleVersePointer> GetAllVerses()
        {
            var result = new List<SimpleVersePointer>();

            result.AddRange(this.VerseNumber.GetAllVerses().ConvertAll(v =>
            {
                var verse = (SimpleVersePointer)this.Clone();
                verse.VerseNumber = new VerseNumber(v);
                return verse;
            }));

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
        public SimpleVerse(SimpleVersePointer versePointer, string verseNumberString, string verseContent)
            : base(versePointer.BookIndex, versePointer.Chapter, versePointer.VerseNumber)
        {
            this.VerseContent = verseContent;

            if (!string.IsNullOrEmpty(verseNumberString))
                this.VerseNumberString = verseNumberString;
            else
                this.VerseNumberString = versePointer.VerseNumber.ToString();            
        }

        public override object Clone()
        {
            var result = new SimpleVerse(this, this.VerseNumberString, this.VerseContent);
            CopyPropertiesTo(result);

            return result;            
        }
    }

    [Serializable]
    public class VersePointer: IComparable<VersePointer>
    {      
        public BibleBookInfo Book { get; set; }
        public int? Chapter { get; set; }
        public int? Verse { get; set; }

        /// <summary>
        /// первоначально переданная строка в конструктор
        /// </summary>
        public string OriginalVerseName { get; set; }   
        public string OriginalBookName { get; set; }        

        public bool WasChangedVerseAsOneChapteredBook { get; set; }

        /// <summary>
        /// родительская ссылка. Например если мы имеем дело со стихом диапазона, то здесь хранится стих, являющийся диапазоном
        /// </summary>
        public VersePointer ParentVersePointer { get; set; }
        

        private VerseNumber? _verseNumber;
        /// <summary>
        /// Стих, который был передан изначально
        /// </summary>
        public VerseNumber VerseNumber
        {
            get
            {
                if (_verseNumber.HasValue)
                    return _verseNumber.Value;

                return new VerseNumber(this.Verse.GetValueOrDefault(), this.TopChapter.HasValue ? null : this.TopVerse);
            }
            set
            {
                _verseNumber = value;
            }
        }        

        public override string ToString()
        {
            if (!string.IsNullOrEmpty(this.OriginalVerseName))
                return this.OriginalVerseName;

            return base.ToString();
        }

        public string ToFirstVerseString()
        {
            return string.Format("{0} {1}:{2}", Book != null ? Book.Name : string.Empty, Chapter, Verse.GetValueOrDefault());
        }

        /// <summary>
        /// Если стих является subverse и стих является первым в родительском стихе, то возвращает true, иначе - false 
        /// </summary>
        public bool IsFirstVerseInParentVerse()
        {
            if (ParentVersePointer != null)
            {
                return this.Verse == ParentVersePointer.Verse;
            }

            return false;
        }


        /// <summary>
        /// Новый термин: MultiVerseString - строка в стихе после названия книги. (*| 5:6, :6, :6-7, 5-6...)
        /// </summary>
        /// <returns></returns>
        public string GetFullMultiVerseString()
        {
            if (IsMultiVerse)
            {
                if (TopChapter != null && TopVerse != null)
                    return string.Format("{0}:{1}-{2}:{3}", Chapter, Verse, TopChapter, TopVerse);
                else if (TopChapter != null && IsChapter)
                    return string.Format("{0}-{1}", Chapter, TopChapter);
                else
                    return string.Format("{0}:{1}", Chapter, VerseNumber);
            }
            else
            {
                if (IsChapter)
                    return string.Format("{0}", Chapter);
                else
                    return string.Format("{0}:{1}", Chapter, VerseNumber);
            }            
        }

        /// <summary>
        /// Возвращает "лёгкую" версию мульти строки стиха. То есть если возможно - без главы.
        /// </summary>
        /// <returns></returns>
        public string GetLightMultiVerseString()
        {
            if (IsMultiVerse)
            {
                if (TopChapter != null && TopVerse != null)
                    return string.Format("{0}:{1}-{2}:{3}", Chapter, Verse, TopChapter, TopVerse);
                else if (TopChapter != null && IsChapter)
                    return string.Format("{0}-{1}", Chapter, TopChapter);
                else
                    return string.Format(":{0}-{1}", Verse, TopVerse);
            }
            else
            {
                if (IsChapter)
                    return string.Format("{0}", Chapter);
                else
                    return string.Format(":{0}", Verse);
            }            
        }

        public static VersePointer GetChapterVersePointer(string bookName, int chapter)
        {
            return GetChapterVersePointer(string.Format("{0} {1}", bookName, chapter));
        }

        public static VersePointer GetChapterVersePointer(string chapterName)
        {
            return new VersePointer(chapterName);
        }

        public string GetFriendlyFullVerseName()
        {
            return string.Format("{0} {1}", Book.FriendlyShortName, GetFullMultiVerseString());            
        }

        public string ChapterName
        {
            get
            {                
                return string.Format("{0} {1}", Book != null ? Book.Name : string.Empty, Chapter);
            }
        }

        public VersePointer(VersePointer chapterPointer, int verse)
            : this(chapterPointer, verse, null)
        {

        }

        public VersePointer(VersePointer chapterPointer, VerseNumber verseNumber)
            : this(chapterPointer, verseNumber.Verse, verseNumber.TopVerse)
        {

        }

        public VersePointer(VersePointer chapterPointer, int verse, int? topVerse)
            : this(chapterPointer.OriginalBookName, chapterPointer.Chapter.Value, verse, topVerse)
        {

        }

        public VersePointer(string bookName, int chapter)
            : this(bookName, chapter, null)
        {

        }

        public VersePointer(string bookName, int chapter, int? verse)
            : this(bookName, chapter, verse, null)
        {

        }

        public VersePointer(string bookName, int chapter, int? verse, int? topVerse)
            : this(string.Format("{0} {1}{2}{3}", 
                                    bookName, 
                                    chapter, 
                                    verse.HasValue ? ":" + verse : string.Empty, 
                                    topVerse.HasValue ? "-" + topVerse : string.Empty))
        {

        }

        public VersePointer(string s)
        {
            if (!string.IsNullOrEmpty(s))
            {
                this.OriginalVerseName = s;

                s = s.ToLowerInvariant();

                s = TrimLocation(s);

                int verseIndex;
                int? verse = StringUtils.GetStringLastNumber(s, out verseIndex);

                if (verse.HasValue)
                {
                    Verse = verse.Value;

                    int i = s.LastIndexOf(Verse.ToString());

                    s = s.Substring(0, i);

                    int chapterIndex;
                    int? chapter = StringUtils.GetStringLastNumber(s, out chapterIndex);

                    if (chapter.HasValue && chapterIndex > 1         // чтобы уберечь от строк типа "1Кор" - то есть считаем, что глава идёт хотя бы с третьего символа
                        && (verseIndex - chapterIndex == verse.Value.ToString().Length + 1))   // чтобы уберечь от строк типа "Во 2Кор 5 и Рим 14"
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
            if (Book == null)
                throw new InvalidOperationException(string.Format("Book is null for {0}", this.OriginalVerseName));

            return new SimpleVersePointer(Book.Index, Chapter.GetValueOrDefault(0), VerseNumber);
        }

        private void ConvertToBaseVerse(string moduleName)
        {
            if (IsValid)
            {
                var parallelVersePointer = BibleParallelTranslationConnectorManager.GetParallelVersePointer(
                                                ToSimpleVersePointer(), moduleName, SettingsManager.Instance.ModuleShortName);

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
                                                moduleName, SettingsManager.Instance.ModuleShortName);

                    if (TopChapter.HasValue)
                        _topChapter = parallelVersePointer.Chapter;
                    if (TopVerse.HasValue)
                        _topVerse = parallelVersePointer.Verse;
                }
            }
        }

        /// <summary>
        /// Это не посто глава, а все стихи этой главы, и мы их трансформировали в главу 
        /// </summary>
        public bool GroupedVerses { get; set; } 

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
        /// Принадлежит ли указанный стих текущему диапазону стихов. Работает только если текущий стих IsMultiVerse, а передаваемый стих - не IsMultiVerse
        /// </summary>
        /// <param name="verse"></param>
        /// <returns></returns>
        public bool IsInVerseRange(VersePointer vp)
        {
            if (!this.IsMultiVerse)
                return false;

            if (this.TopChapter.HasValue)
            {
                return (this.Chapter < vp.Chapter && vp.Chapter < this.TopChapter)
                    || (this.Chapter == vp.Chapter && this.Verse <= vp.Verse.GetValueOrDefault(0))
                    || (this.TopChapter == vp.Chapter && vp.Verse.GetValueOrDefault(0) <= this.TopVerse);
            }
            else
                return this.Verse <= vp.Verse.GetValueOrDefault(0) && vp.Verse.GetValueOrDefault(0) <= this.TopVerse;
        }


        private string TrimLocation(string s)
        {            
            int indexOfDash = s.LastIndexOfAny(VerseNumber.Dashes);

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
                    int indexOfDelimiter = topStructure.IndexOf(VerseRecognitionManager.DefaultChapterVerseDelimiter);  // даже если используем запятую в качестве разделителя, сюда приходит уже DefaultChapterVerseDelimiter
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
                var delimiters = new char[] { '.', ',', ':' };
                int iDel = s.IndexOfAny(delimiters);
                if (char.IsLetter(StringUtils.GetChar(s, iDel - 1)))
                    iDel = s.IndexOfAny(delimiters, iDel + 1);

                if (iDel != -1)
                {
                    if (iComma > iDel) // ситуация типа "2Тим 1:13,14"
                        s = s.Substring(0, iComma);
                }
            }

            return s;
        }

        public void ChangeVerseAsOneChapteredBook()
        {
            Verse = Chapter;
            _topVerse = TopChapter;

            Chapter = 1;
            _topChapter = null;

            WasChangedVerseAsOneChapteredBook = true;
        }

        private static string TrimBookName(string bookName, out bool endsWithDot)
        {
            string result = bookName.Trim(' ');
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

        public bool? IsExisting
        {
            get
            {
                if (!SettingsManager.Instance.CanUseBibleContent)
                    return null;

                VerseNumber vn;
                var svp = ToSimpleVersePointer();
                var result = SettingsManager.Instance.CurrentBibleContentCached.VerseExists(svp, SettingsManager.Instance.ModuleShortName, out vn);

                if (!result && IsChapter && SettingsManager.Instance.CurrentBibleContentCached.BookHasOnlyOneChapter(svp))                
                    result = SettingsManager.Instance.CurrentBibleContentCached.VerseExists(svp.GetChangedVerseAsOneChapteredBook(), SettingsManager.Instance.ModuleShortName, out vn);                

                return result;
            }
        }

        private static BibleBookInfo GetBibleBook(string s, bool endsWithDot, out string moduleName)
        {   
            return SettingsManager.Instance.CurrentModuleCached.GetBibleBook(s, endsWithDot, out moduleName);
        }

        public VersePointer GetChapterPointer()
        {
            return new VersePointer(this.OriginalBookName, this.Chapter.Value, 0);
        }


        public GetAllIncludedVersesResult GetAllVerses(ref Application oneNoteApp, GetAllIncludedVersesArgs args)
        {
            var result = new GetAllIncludedVersesResult();

            var firstVerse = this.IsChapter ? new VersePointer(this.OriginalBookName, this.Chapter.Value) : new VersePointer(this.OriginalBookName, this.Chapter.Value, this.Verse.Value);
            firstVerse.ParentVersePointer = this;

            var allVersesExceptFirst = GetAllIncludedVersesExceptFirst(ref oneNoteApp, args);
            result.VersesCount = allVersesExceptFirst.VersesCount + 1;
            
            if (!allVersesExceptFirst.NotNeedFirstVerse.GetValueOrDefault(false))
                result.Verses.Add(firstVerse);

            result.Verses.AddRange(allVersesExceptFirst.Verses);
            
            return result;
        }

        /// <summary>
        /// возвращает список всех вложенных стихов (за исключением первого) если Multiverse. 
        /// </summary>
        /// <returns></returns>
        public GetAllIncludedVersesResult GetAllIncludedVersesExceptFirst(ref Application oneNoteApp, GetAllIncludedVersesArgs args)
        {
            var result = new GetAllIncludedVersesResult();
            if ((IsValid || args.Force) && IsMultiVerse)
            {                
                if (TopChapter != null && TopVerse != null && !args.SearchOnlyForFirstChapter)
                {
                    result = GetAllIncludedVersesExceptFirstFromMultiChaptersAndVerses(ref oneNoteApp, args);                   
                }
                else if (TopChapter != null && IsChapter)
                {
                    result = GetAllIncludedVersesExceptFirstFromMultiChapters(ref oneNoteApp, args);
                }
                else if (!TopChapter.HasValue)  // в будущем, когда у нас будут выделяться все стихи из диапазона на странице Библии, надо будет переделать этот метод, чтобы если TopChapter.HasValue - то он доставал все стихи текущей главы (потому что до сюда он может дойти (если указана TopChapter) только если указан args.SearchOnlyForFirstChapter)
                {
                    result = GetAllIncludedVersesExceptFirstFromMultiVersesOfFirstChapterOnly(ref oneNoteApp, args);
                }
            }

            return result;
        }

        private GetAllIncludedVersesResult GetAllIncludedVersesExceptFirstFromMultiVersesOfFirstChapterOnly(ref Application oneNoteApp, GetAllIncludedVersesArgs args)
        {
            var result = new GetAllIncludedVersesResult();

            var loadChapterOnly = false;

            if (args.TryToGroupVersesInChapters)
            {
                if (Verse.GetValueOrDefault(0) == 1)
                {
                    var chapterVersesCount = HierarchySearchManager.GetChapterVersesCount(
                                                ref oneNoteApp, args.BibleNotebookId,
                                                VersePointer.GetChapterVersePointer(this.OriginalBookName, this.Chapter.Value), null, null)
                                                .GetValueOrDefault(0);

                    if (TopVerse >= chapterVersesCount)
                        loadChapterOnly = true;
                }
            }

            if (loadChapterOnly)
            {
                var vp = this.GetChapterPointer();
                vp.ParentVersePointer = this;
                vp.GroupedVerses = true;
                result.Verses.Add(vp);
                result.NotNeedFirstVerse = true;
            }
            else
            {
                for (int verseIndex = Verse.GetValueOrDefault(0) + 1; verseIndex <= TopVerse; verseIndex++)
                {
                    VersePointer vp = new VersePointer(this.OriginalBookName, this.Chapter.Value, verseIndex);
                    vp.ParentVersePointer = this;
                    result.Verses.Add(vp);
                }
            }

            result.VersesCount = TopVerse.Value - Verse.GetValueOrDefault(0);

            return result;
        }

        private GetAllIncludedVersesResult GetAllIncludedVersesExceptFirstFromMultiChapters(ref Application oneNoteApp, GetAllIncludedVersesArgs args)
        {
            var result = new GetAllIncludedVersesResult();

            for (int chapterIndex = Chapter.Value + 1; chapterIndex <= TopChapter; chapterIndex++)
            {
                VersePointer vp = VersePointer.GetChapterVersePointer(this.OriginalBookName, chapterIndex);
                vp.ParentVersePointer = this;
                result.Verses.Add(vp);
            }

            result.VersesCount = result.Verses.Count;

            return result;
        }

        private GetAllIncludedVersesResult GetAllIncludedVersesExceptFirstFromMultiChaptersAndVerses(ref Application oneNoteApp, GetAllIncludedVersesArgs args)
        {
            var result = new GetAllIncludedVersesResult();
            var versesCount = 0;
            for (int chapterIndex = Chapter.Value; chapterIndex <= TopChapter; chapterIndex++)
            {
                var wasLoadedChapterVersesCount = false;
                int topVerse, startVerseIndex;
                if (chapterIndex == TopChapter)
                    topVerse = TopVerse.Value;
                else
                {
                    wasLoadedChapterVersesCount = true;
                    topVerse = HierarchySearchManager.GetChapterVersesCount(
                                    ref oneNoteApp, args.BibleNotebookId,
                                    VersePointer.GetChapterVersePointer(this.OriginalBookName, chapterIndex), null, null)
                                    .GetValueOrDefault(0);
                }

                if (chapterIndex == Chapter)
                    startVerseIndex = Verse.Value + 1;
                else
                    startVerseIndex = 1;

                var loadChapterOnly = false;

                if (args.TryToGroupVersesInChapters)
                {
                    if (startVerseIndex == 1 || (chapterIndex == Chapter && Verse.Value == 1))
                    {
                        if (!wasLoadedChapterVersesCount)
                        {
                            var chapterVersesCount = HierarchySearchManager.GetChapterVersesCount(
                                                        ref oneNoteApp, args.BibleNotebookId,
                                                        VersePointer.GetChapterVersePointer(this.OriginalBookName, chapterIndex), null, null)
                                                        .GetValueOrDefault(0);

                            if (topVerse >= chapterVersesCount)
                                loadChapterOnly = true;
                        }
                        else
                            loadChapterOnly = true;
                    }
                }

                if (loadChapterOnly)
                {
                    VersePointer vp = VersePointer.GetChapterVersePointer(this.OriginalBookName, chapterIndex);
                    vp.ParentVersePointer = this;
                    vp.GroupedVerses = true;
                    result.Verses.Add(vp);
                    versesCount += topVerse - startVerseIndex + 1;

                    if (chapterIndex == Chapter.Value)
                        result.NotNeedFirstVerse = true;
                }
                else
                {
                    for (int verseIndex = startVerseIndex; verseIndex <= topVerse; verseIndex++)
                    {
                        VersePointer vp = new VersePointer(this.OriginalBookName, chapterIndex, verseIndex);
                        vp.ParentVersePointer = this;
                        result.Verses.Add(vp);
                        versesCount++;
                    }
                }
            }

            result.VersesCount = versesCount;

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

        public int CompareTo(VersePointer other)
        {
            if (other == null)
                return 1;

            return this.Verse.GetValueOrDefault(0).CompareTo(other.Verse.GetValueOrDefault(0));
        }
    }
}
