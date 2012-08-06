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


            foreach (int bookIndex in baseTranslationDifferencesEx.BibleVersesDifferences.Keys)
            {
                var bookVersePointersComparisonTables = new SimpleVersePointersComparisonTable();
                result.Add(bookIndex, bookVersePointersComparisonTables);

                if (parallelTranslationDifferencesEx.BibleVersesDifferences.ContainsKey(bookIndex))
                {
                    var baseBookVerses = baseTranslationDifferencesEx.BibleVersesDifferences[bookIndex];
                    var parallelBookVerses = parallelTranslationDifferencesEx.BibleVersesDifferences[bookIndex];

                    foreach (var parallelVerseKey in parallelBookVerses.Keys)
                    {
                        var parallelVerseValue = parallelBookVerses[parallelVerseKey];

                        if (baseBookVerses.ContainsKey(parallelVerseKey))
                            bookVersePointersComparisonTables.Add(baseBookVerses[parallelVerseKey], parallelVerseValue);
                        else
                            bookVersePointersComparisonTables.Add(parallelVerseKey, parallelVerseValue);
                    }
                }
            }


            return result;
        }
    }
}
