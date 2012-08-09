using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BibleCommon.Common
{
    public class ComparisonVersesInfo: List<SimpleVersePointer>
    {
        public bool Strict { get; set; }
        public BibleBookDifference.VerseAlign Align { get; set; }
        public int? ValueVerseCount { get; set; }

        public ComparisonVersesInfo()
        {
        }

        public ComparisonVersesInfo(List<SimpleVersePointer> verses)
            : base(verses)
        {
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
            var baseVersesFormula = new BibleTranslationDifferencesBaseVersesFormula(bookIndex, bookDifference.BaseVerses);
            var parallelVersesFormula = new BibleTranslationDifferencesParallelVersesFormula(bookDifference.ParallelVerses, baseVersesFormula, 
                bookDifference.Align, bookDifference.Strict);            

            foreach (var verse in baseVersesFormula.GetAllVerses())
            {
                var parallelVerses = parallelVersesFormula.GetParallelVerses(verse);
                parallelVerses.Align = bookDifference.Align;
                parallelVerses.Strict = bookDifference.Strict;
                parallelVerses.ValueVerseCount = bookDifference.ValueVerseCount;

                BibleVersesDifferences[bookIndex].Add(verse, parallelVerses);
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
