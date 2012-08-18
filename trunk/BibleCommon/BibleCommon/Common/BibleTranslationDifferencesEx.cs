using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BibleCommon.Common
{
    public class ComparisonVersesInfo: List<SimpleVersePointer>
    {
        public ComparisonVersesInfo()
        {
        }

        public ComparisonVersesInfo(List<SimpleVersePointer> verses)
            : base(verses)
        {
        }

        public static ComparisonVersesInfo FromVersePointer(SimpleVersePointer versePointer)
        {
            return new ComparisonVersesInfo() { versePointer };     
        }
    }

    public class SimpleVersePointersComparisonTable : Dictionary<SimpleVersePointer, ComparisonVersesInfo>
    {
        public new void Add(SimpleVersePointer key, ComparisonVersesInfo value)
        {
            if (!this.ContainsKey(key))
                base.Add(key, value);
            else
                base[key].AddRange(value);
        }        
    }

    public class BibleTranslationDifferencesEx
    {
        public Dictionary<int, SimpleVersePointersComparisonTable> BibleVersesDifferences { get; set; }

        public BibleTranslationDifferencesEx(BibleTranslationDifferences translationDifferences)
        {
            BibleVersesDifferences = new Dictionary<int, SimpleVersePointersComparisonTable>();

            foreach (var bookDifferences in translationDifferences.BookDifferences)
            {
                BibleVersesDifferences.Add(bookDifferences.BookIndex, new SimpleVersePointersComparisonTable());

                foreach (var bookDifference in bookDifferences.Differences)
                {
                    ProcessBookDifference(bookDifferences.BookIndex, bookDifference);
                }
            }
        }        

        private void ProcessBookDifference(int bookIndex, BibleBookDifference bookDifference)
        {
            int? valueVersesCount = string.IsNullOrEmpty(bookDifference.ValueVersesCount) ? (int?)null : int.Parse(bookDifference.ValueVersesCount);

            var baseVersesFormula = new BibleTranslationDifferencesBaseVersesFormula(bookIndex, bookDifference.BaseVerses, bookDifference.CorrespondenceType);
            var parallelVersesFormula = new BibleTranslationDifferencesParallelVersesFormula(bookDifference.ParallelVerses, baseVersesFormula,
                bookDifference.CorrespondenceType, valueVersesCount);

            SimpleVersePointer prevVerse = null;
            foreach (var verse in baseVersesFormula.GetAllVerses())
            {
                var parallelVerses = new ComparisonVersesInfo(parallelVersesFormula.GetParallelVerses(verse, prevVerse));                                

                BibleVersesDifferences[bookIndex].Add(verse, parallelVerses);

                prevVerse = parallelVerses.Last();
            }
        }        

        public SimpleVersePointersComparisonTable GetBibleVersesDifferences(int bookIndex)
        {
            if (this.BibleVersesDifferences.ContainsKey(bookIndex))
                return this.BibleVersesDifferences[bookIndex];

            return null;
        }
    }
}
