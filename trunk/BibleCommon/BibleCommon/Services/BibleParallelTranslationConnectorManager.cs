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
                        var baseVerses = baseBookVerses[baseVerseKey];
                        int? versePartIndex = baseVerses.Count > 1 ? (int?)0 : null;
                        foreach(var baseVersePointer in baseVerses)
                        {
                            var parallelVerses = ComparisonVersesInfo.FromVersePointer(
                                new SimpleVersePointer(baseVerseKey) 
                                { 
                                    PartIndex = versePartIndex.HasValue ? versePartIndex++ : null,
                                    IsEmpty = baseVersePointer.IsApocrypha
                                },
                                baseVerses.Strict);                            

                            bookVersePointersComparisonTables.Add(baseVersePointer, parallelVerses); 
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
                var notApocryphaBaseVerses = baseVerses.Where(v => !v.IsApocrypha);
                var notApocryphaParallelVerses = parallelVerses.Where(v => !v.IsApocrypha);

                bool isPartVersePointer = notApocryphaParallelVerses.Count() < notApocryphaBaseVerses.Count();


                int parallelVerseIndex = 0;
                int partIndex = 0;
                for (int baseVerseIndex = 0; baseVerseIndex < baseVerses.Count; baseVerseIndex++)
                {
                    var baseVerse = baseVerses[baseVerseIndex];
                    ComparisonVersesInfo parallelVersesInfo = new ComparisonVersesInfo();
                    //parallelVersesInfo.Strict = parallelVerses.Strict;

                    if (baseVerse.IsApocrypha)
                    { 
                        int lastParallelVerseIndex = parallelVerseIndex;
                        for (int i = parallelVerseIndex; i < parallelVerses.Count; i++)
                        {
                            lastParallelVerseIndex = i + 1;
                            var parallelVerseToAdd = (SimpleVersePointer)parallelVerses[i].Clone();
                            if (!parallelVerseToAdd.IsApocrypha)
                            {
                                ???
                            }                            
                        }
                        parallelVerseIndex = lastParallelVerseIndex;
                    }
                    else
                    {
                        int lastParallelVerseIndex = parallelVerseIndex;
                        for (int i = parallelVerseIndex; i < parallelVerses.Count; i++)
                        {
                            lastParallelVerseIndex = i + 1;
                            var parallelVerseToAdd = (SimpleVersePointer)parallelVerses[i].Clone();

                            if (!parallelVerseToAdd.IsApocrypha)
                            {
                                parallelVerseToAdd.PartIndex = isPartVersePointer ? (int?)partIndex++ : null;
                                parallelVersesInfo.Add(parallelVerseToAdd);
                                break;
                            }
                            else
                            {
                                parallelVersesInfo.Add(parallelVerseToAdd);
                            }                            
                        }
                        parallelVerseIndex = lastParallelVerseIndex;
                    }

                    bookVersePointersComparisonTables.Add(baseVerse, parallelVersesInfo);
                }

                



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
