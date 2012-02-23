﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections.Specialized;
using BibleCommon.Consts;
using BibleCommon.Helpers;

namespace BibleCommon.Common
{
    public class VersePointer
    {      
        public string BookName { get; set; }
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

        public static VersePointer GetChapterVersePointer(string chapterName)
        {
            return new VersePointer(string.Format("{0}:0", chapterName));
        }

        public string FriendlyVerseName
        {
            get
            {
                return string.Format("{0}:{1}", FriendlyChapterName, Verse);
            }
        }

        public string FriendlyChapterName
        {
            get
            {
                if (BookName.Length < 4)
                    throw new InvalidOperationException(string.Format("BookName can not be less than 4 symbols: '{0}'.", BookName));

                return string.Format("{0} {1}", BookName.Substring(4), Chapter);
            }
        }

        public string ChapterName
        {
            get
            {
                return string.Format("{0} {1}", BookName, Chapter);
            }
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

                    int? chapter = StringUtils.GetStringLastNumber(s);

                    if (chapter.HasValue)
                    {
                        Chapter = chapter.Value;

                        i = s.LastIndexOf(Chapter.ToString());

                        s = s.Substring(0, i);
                    }

                    if (TopVerse <= Verse)
                    {
                        _isMultiVerse = false;
                        _topVerse = null;
                    }
                }

                OriginalBookName = TrimBookName(s);
                BookName = GetFullBookName(OriginalBookName);
            }
        }

        public bool IsChapter
        {
            get
            {
                return this.Verse.GetValueOrDefault(0) == 0;
            }
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
                if (int.TryParse(s.Substring(indexOfDash + 1), out tempTopVerse))
                {
                    _topVerse = tempTopVerse;
                    _isMultiVerse = true;
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

        private static string TrimBookName(string bookName)
        {
            for (int index = bookName.Length - 1; index >= 0; index--)
            {
                char c = bookName[index];

                if (StringUtils.IsCharAlphabetical(c) || StringUtils.IsDigit(c))
                    break;
                else if (c == ' ' || c == '.')
                    bookName = bookName.Remove(bookName.Length - 1);
            }

            if (bookName.IndexOf("  ") != -1)  // двойной пробел
            {
                bookName = bookName.Replace("  ", " ");
            }

            return bookName.Trim();
        }

        public bool IsValid
        {
            get
            {
                return !string.IsNullOrEmpty(BookName)
                    && Chapter.HasValue
                    && Verse.HasValue;
            }
        }

        private static string GetFullBookName(string s)
        {
            s = s.ToLowerInvariant();

            foreach (string key in Constants.BookNames.Keys)
            {
                if (Constants.BookNames[key].Contains(s))
                    return key;  // возвращаться должно в правильном Case!
            }

            return string.Empty;
        }

        /// <summary>
        /// возвращает список всех вложенных стихов (за исключением первого) если Multiverse. 
        /// </summary>
        /// <returns></returns>
        public List<VersePointer> GetAllIncludedVersesExceptFirst()
        {
            List<VersePointer> result = new List<VersePointer>();
            if (IsValid && IsMultiVerse && !IsChapter)
            {
                for (int i = Verse.GetValueOrDefault(0) + 1; i <= TopVerse; i++)
                {
                    VersePointer vp = new VersePointer(this.OriginalBookName, this.Chapter.Value, i);
                    vp.ParentVersePointer = this;
                    result.Add(vp);                   
                }
            }

            return result;
        }

        public override bool Equals(object obj)
        {
            VersePointer otherVp = (VersePointer)obj;
            return this.BookName == otherVp.BookName
                && this.Chapter == otherVp.Chapter
                && this.Verse == otherVp.Verse;                
        }

        public override int GetHashCode()
        {
            return this.BookName.GetHashCode() ^ this.Chapter.GetHashCode() ^ this.Verse.GetHashCode();
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
