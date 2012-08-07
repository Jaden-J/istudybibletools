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
                        bookVersePointersComparisonTables.Add(baseBookVerses[baseVerseKey], parallelBookVerses[baseVerseKey]);
                    else
                        bookVersePointersComparisonTables.Add(baseBookVerses[baseVerseKey], baseVerseKey);
                }
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

                var baseBookVerses = baseTranslationDifferencesEx.GetBibleVersesDifferences(bookIndex);
                var parallelBookVerses = parallelTranslationDifferencesEx.BibleVersesDifferences[bookIndex];

                foreach (var parallelVerseKey in parallelBookVerses.Keys)
                {
                    if (baseBookVerses != null && baseBookVerses.ContainsKey(parallelVerseKey))
                        bookVersePointersComparisonTables.Add(baseBookVerses[parallelVerseKey], parallelBookVerses[parallelVerseKey]);
                    else
                        bookVersePointersComparisonTables.Add(parallelVerseKey, parallelBookVerses[parallelVerseKey]);

                }
            }
        }
    }
}
