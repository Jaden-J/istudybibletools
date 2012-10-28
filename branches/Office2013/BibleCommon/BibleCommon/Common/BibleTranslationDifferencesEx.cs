﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BibleCommon.Services;

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

        private Dictionary<SimpleVersePointer, SimpleVersePointer> _keys;
        public SimpleVersePointer GetOriginalKey(SimpleVersePointer key)
        {
            if (_keys == null)
            {
                _keys = new Dictionary<SimpleVersePointer, SimpleVersePointer>();
                foreach (var k in this.Keys)
                {
                    _keys.Add(k, k);
                }                
            }

            if (_keys.ContainsKey(key))
                return _keys[key];

            return null;
        }
    }

    public class BibleTranslationDifferencesEx
    {
        public ParallelBibleInfo BibleVersesDifferences { get; set; }

        public BibleTranslationDifferencesEx(BibleTranslationDifferences translationDifferences)
        {
            BibleVersesDifferences = new ParallelBibleInfo();

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

            var baseVersesFormula = new BibleTranslationDifferencesBaseVersesFormula(bookIndex, bookDifference.BaseVerses, bookDifference.ParallelVerses, 
                                                    bookDifference.CorrespondenceType, bookDifference.SkipCheck);
            var parallelVersesFormula = new BibleTranslationDifferencesParallelVersesFormula(bookDifference.ParallelVerses, baseVersesFormula,
                bookDifference.CorrespondenceType, valueVersesCount, bookDifference.SkipCheck);

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