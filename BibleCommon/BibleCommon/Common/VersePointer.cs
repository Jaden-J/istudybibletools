using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections.Specialized;
using BibleCommon.Consts;
using BibleCommon.Helpers;
using BibleCommon.Services;
using Microsoft.Office.Interop.OneNote;

namespace BibleCommon.Common
{
    public class SimpleVersePointer: ICloneable
    {
        public int BookIndex { get; set; }
        public int Chapter { get; set; }
        public int Verse { get; set; }
        public int? PartIndex { get; set; }
        public int? TopVerse { get; set; }
        public bool IsEmpty { get; set; }
        public bool IsApocrypha { get; set; }
        //public SimpleVersePointer BaseVersePointer { get; set; }  // если IsApocrypha - стих, к которому привязан данный стих

        public SimpleVersePointer(SimpleVersePointer verse)
            : this(verse.BookIndex, verse.Chapter, verse.Verse)
        {

        }

        public SimpleVersePointer(int bookIndex, int chapter, int verse)
        {
            this.BookIndex = bookIndex;
            this.Chapter = chapter;
            this.Verse = verse;
        }

        public override bool Equals(object obj)
        {
            if (obj == null)
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
            var result = string.Format("{0} {1}:{2}", BookIndex, Chapter, Verse);

            if (PartIndex.HasValue)
                result += string.Format("({0})", PartIndex);

            if (IsEmpty)
                result += "(empty)";

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
            verse.TopVerse = this.TopVerse;
        }
    }

    public class SimpleVerse : SimpleVersePointer
    {
        public string VerseContent { get; set; }

        public SimpleVerse(SimpleVersePointer versePointer, string verseContent)
            : base(versePointer.BookIndex, versePointer.Chapter, versePointer.Verse)
        {
            this.VerseContent = verseContent;
        }

        public override object Clone()
        {
            var result = new SimpleVerse(this, this.VerseContent);
            CopyProperties(result);

            return result;            
        }
    }

    public class VersePointer
    {      
        public BibleBookInfo Book { get; set; }
        public int? Chapter { get; set; }
        public int? Verse { get; set; }
        public string OriginalVerseName { get; set; }   // первоначально переданная строка в конструктор
        public string OriginalBookName { get; set; }

        public VersePointer ParentVersePointer { get; set; } // родительская ссылка. Например если мы имеем дело со стихом диапазона, то здесь хранится стих, являющийся диапазоном

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

        public VersePointer(string bookName, int chapter, int verse)
            : this(string.Format("{0} {1}:{2}", bookName, chapter, verse))
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
                OriginalBookName = TrimBookName(s, out endsWithDot);
                Book = GetBibleBook(OriginalBookName, endsWithDot);
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

        private static BibleBookInfo GetBibleBook(string s, bool endsWithDot)
        {
            return SettingsManager.Instance.CurrentModule.GetBibleBook(s, endsWithDot);
        }

        public VersePointer GetChapterPointer()
        {
            return new VersePointer(this.OriginalBookName, this.Chapter.Value, 0);
        }

        /// <summary>
        /// возвращает список всех вложенных стихов (за исключением первого) если Multiverse. 
        /// </summary>
        /// <returns></returns>
        public List<VersePointer> GetAllIncludedVersesExceptFirst(Application oneNoteApp, string bibleNotebookId, bool force = false)
        {
            List<VersePointer> result = new List<VersePointer>();
            if ((IsValid || force) && IsMultiVerse)
            {
                if (TopChapter != null && TopVerse != null)
                {
                    for (int chapterIndex = Chapter.Value; chapterIndex <= TopChapter; chapterIndex++)
                    {
                        int topVerse, startVerseIndex;
                        if (chapterIndex == TopChapter)
                            topVerse = TopVerse.Value;
                        else
                            topVerse = HierarchySearchManager.GetChapterVersesCount(
                                            oneNoteApp, bibleNotebookId, 
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
            var result =  this.Chapter.GetHashCode() ^ this.Verse.GetHashCode();
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
            if (((object)vp1) == null && ((object)vp2) == null)
                return false;

            if (((object)vp1) == null)
                return true;

            if (((object)vp2) == null)
                return true;

            return !vp1.Equals(vp2);
        }
    }
}
