using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BibleCommon.Common;

namespace BibleCommon.Services
{
    public static class BibleParallelTranslationConnectorManager
    {
        public static Dictionary<int, SimpleVersePointersComparisonTable> ConnectBibleTranslations(BibleTranslationDifferences baseBookTranslationDifferences,
            BibleTranslationDifferences parallelBookTranslationDifferences)
        {
            var result = new Dictionary<int, SimpleVersePointersComparisonTable>();

            var baseTranslationDifferencesEx = new BibleTranslationDifferencesEx(baseBookTranslationDifferences);
            var parallelTranslationDifferencesEx = new BibleTranslationDifferencesEx(parallelBookTranslationDifferences);


            ProcessForBaseBookVerses(baseTranslationDifferencesEx, parallelTranslationDifferencesEx, result);
            ProcessForParallelBookVerses(baseTranslationDifferencesEx, parallelTranslationDifferencesEx, result);
          
            return result;
        }

       

        private static void ProcessForBaseBookVerses(BibleTranslationDifferencesEx baseTranslationDifferencesEx, BibleTranslationDifferencesEx parallelTranslationDifferencesEx,
            Dictionary<int, SimpleVersePointersComparisonTable> result)
        {
            foreach (int bookIndex in baseTranslationDifferencesEx.BibleVersesDifferences.Keys)
            {
                var bookVersePointersComparisonTables = new SimpleVersePointersComparisonTable();
                result.Add(bookIndex, bookVersePointersComparisonTables);

                var baseBookVerses = baseTranslationDifferencesEx.BibleVersesDifferences[bookIndex];
                var parallelBookVerses = parallelTranslationDifferencesEx.GetBibleVersesDifferences(bookIndex);

                foreach (var baseVerseKey in baseBookVerses.Keys)
                {
                    if (parallelBookVerses != null && parallelBookVerses.ContainsKey(baseVerseKey))
                    {
                        var baseVerses = baseBookVerses[baseVerseKey];
                        var parallelVerses = parallelBookVerses[baseVerseKey];

                        JoinBaseAndParallelVerses(baseVerseKey, baseVerses, parallelVerses, bookVersePointersComparisonTables);
                    }
                    else
                    {   
                        int versePartIndex = 0;
                        foreach(var baseVersePointer in baseBookVerses[baseVerseKey])
                        {
                            bookVersePointersComparisonTables.Add(baseVersePointer, new ComparisonVersesInfo() 
                            { 
                                new SimpleVersePointer(baseVerseKey) { PartIndex = versePartIndex++ }
                            });
                        }
                    }
                }
            }
        }

        private static void JoinBaseAndParallelVerses(SimpleVersePointer versesKey, ComparisonVersesInfo baseVerses, ComparisonVersesInfo parallelVerses,
            SimpleVersePointersComparisonTable bookVersePointersComparisonTables)
        {
            if (baseVerses.Count == 1)
            {
                bookVersePointersComparisonTables.Add(baseVerses[0], parallelVerses);
            }
            else
            {
                throw new NotSupportedException("This case (when baseVerses.Count != 1) is not supported yet.");

                //var baseAlign = baseVerses.Align != BibleBookDifference.VerseAlign.None
                //                        ? baseVerses.Align
                //                        : (versesKey.Verse == 1 ? BibleBookDifference.VerseAlign.Bottom : BibleBookDifference.VerseAlign.Top);
                //var parallelAlign = parallelVerses.Align != BibleBookDifference.VerseAlign.None
                //                        ? parallelVerses.Align
                //                        : (versesKey.Verse == 1 ? BibleBookDifference.VerseAlign.Bottom : BibleBookDifference.VerseAlign.Top);

                //int baseValueVersesCount = baseVerses.ValueVerseCount ?? baseVerses.Count;
                //int parallelValuVersesCount = parallelVerses.ValueVerseCount ?? parallelVerses.Count;

                ////if (baseValueVersesCount != parallelValuVersesCount)
                ////    кидаем warning.

                //if (baseVerses.Count < parallelVerses.Count)
                //{
                //    if (baseAlign == BibleBookDifference.VerseAlign.Top)
                //    {
                //        //for (int 
                //    }
                //}
                //else
                //{

                //}
            }
        }

        private static void ProcessForParallelBookVerses(BibleTranslationDifferencesEx baseTranslationDifferencesEx, BibleTranslationDifferencesEx parallelTranslationDifferencesEx,
           Dictionary<int, SimpleVersePointersComparisonTable> result)
        {
            foreach (int bookIndex in parallelTranslationDifferencesEx.BibleVersesDifferences.Keys)
            {
                SimpleVersePointersComparisonTable bookVersePointersComparisonTables;

                if (!result.ContainsKey(bookIndex))
                {
                    bookVersePointersComparisonTables = new SimpleVersePointersComparisonTable();
                    result.Add(bookIndex, bookVersePointersComparisonTables);
                }
                else
                    bookVersePointersComparisonTables = result[bookIndex];
                
                var parallelBookVerses = parallelTranslationDifferencesEx.BibleVersesDifferences[bookIndex];

                foreach (var parallelVerseKey in parallelBookVerses.Keys)
                {
                    // вариант, когда и там, и там есть мы уже разобрали
                    bookVersePointersComparisonTables.Add(parallelVerseKey, parallelBookVerses[parallelVerseKey]);                    
                }
            }
        }
    }
}
