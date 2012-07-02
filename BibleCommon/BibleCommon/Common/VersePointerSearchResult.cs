using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BibleCommon;
using System.Xml.Linq;
using BibleCommon.Common;
using BibleCommon.Helpers;

namespace BibleCommon.Common
{
    public class VersePointerSearchResult
    {
        public static bool IsChapterAndVerse(SearchResultType resultType)
        {
            return resultType == SearchResultType.ChapterAndVerse
                    || resultType == SearchResultType.ChapterAndVerseAtStartString
                    || resultType == SearchResultType.ChapterAndVerseWithoutBookName;
        }

        public static bool IsVerse(SearchResultType resultType)
        {
            return resultType == SearchResultType.SingleVerseOnly
                    || resultType == SearchResultType.FollowingVerseOnly;
        }

        public static bool IsChapter(SearchResultType resultType)
        {
            return resultType == SearchResultType.ChapterOnly
                    || resultType == SearchResultType.ChapterOnlyAtStartString
                    || resultType == SearchResultType.ExcludableChapter
                    || resultType == SearchResultType.ChapterWithoutBookName;
        }

        public enum SearchResultType
        {
            Nothing = 0,
            ChapterOnly = 1,
            ChapterOnlyAtStartString = 2,
            ExcludableChapter = 3,  // глава, указаная в заголовке в скобках [] (например, [Ин 3]). Для стихов такой главы не создаются ссылки на текущую заметку
            SingleVerseOnly = 4,
            FollowingVerseOnly = 5,   // стих, который следует за запятой
            ChapterAndVerse = 6,
            ChapterAndVerseAtStartString = 7,   // полная ссылка с начала строки
            ChapterAndVerseWithoutBookName = 8,   // например 4:5, а книга берётся из предыдущего результата
            ChapterWithoutBookName = 9,             // например ",4", а книга берётся из предыдущего результата
        }

        public VersePointer VersePointer { get; set; }
        public int VersePointerStartIndex { get; set; }
        public int VersePointerEndIndex { get; set; }
        public int VersePointerHtmlStartIndex { get; set; }
        public int VersePointerHtmlEndIndex { get; set; }
        public string ChapterName { get; set; }
        public string VerseString { get; set; }
        public SearchResultType ResultType { get; set; }
        public XElement TextElement { get; set; }   // где нашли        

        private int? _depthLevel = null;
        public int DepthLevel
        {
            get
            {
                if (TextElement == null)
                    throw new InvalidOperationException("TextElement is null");

                if (!_depthLevel.HasValue)
                {
                    _depthLevel = OneNoteUtils.GetDepthLevel(TextElement);
                }

                return _depthLevel.Value;
            }
        }

        public VersePointerSearchResult()
        {
            this.ResultType = VersePointerSearchResult.SearchResultType.Nothing;
        }
    }
}
