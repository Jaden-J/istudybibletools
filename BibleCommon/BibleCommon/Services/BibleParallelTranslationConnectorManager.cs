﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BibleCommon.Common;

namespace BibleCommon.Services
{
    public static class BibleParallelTranslationConnectorManager
    {
        public static Dictionary<int, SimpleVersePointersComparisonTable> ConnectBibleTranslations(string baseModuleShortName, string parallelModuleShortName,
            BibleTranslationDifferences baseBookTranslationDifferences,
            BibleTranslationDifferences parallelBookTranslationDifferences)
        {
            var result = new Dictionary<int, SimpleVersePointersComparisonTable>();

            if (baseModuleShortName.ToLower() != parallelModuleShortName.ToLower())
            {
                var baseTranslationDifferencesEx = new BibleTranslationDifferencesEx(baseBookTranslationDifferences);
                var parallelTranslationDifferencesEx = new BibleTranslationDifferencesEx(parallelBookTranslationDifferences);

                ProcessForBaseBookVerses(baseTranslationDifferencesEx, parallelTranslationDifferencesEx, result);
                ProcessForParallelBookVerses(baseTranslationDifferencesEx, parallelTranslationDifferencesEx, result);
            }

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
                                });                            

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
                if (parallelVerses.Count == 1 && baseVerses[0].PartIndex.GetValueOrDefault(-1) == parallelVerses[0].PartIndex.GetValueOrDefault(-1))
                {
                    var baseVerseToAdd = (SimpleVersePointer)baseVerses[0].Clone();
                    var parallelVerseToAdd = (SimpleVersePointer)parallelVerses[0].Clone();
                    baseVerseToAdd.PartIndex = null;
                    parallelVerseToAdd.PartIndex = null;
                    if (!bookVersePointersComparisonTables.ContainsKey(baseVerseToAdd))
                        bookVersePointersComparisonTables.Add(baseVerseToAdd, ComparisonVersesInfo.FromVersePointer(parallelVerseToAdd));
                }
                else                
                    bookVersePointersComparisonTables.Add(baseVerses[0], parallelVerses);
            }
            else
            {
                var notApocryphaBaseVerses = baseVerses.Where(v => !v.IsApocrypha);
                var notApocryphaParallelVerses = parallelVerses.Where(v => !v.IsApocrypha);

                bool isPartVersePointer = notApocryphaParallelVerses.Count() < notApocryphaBaseVerses.Count();


                int parallelVerseIndex = 0;                
                List<SimpleVersePointer> prevParallelVerses = new List<SimpleVersePointer>();

                for (int baseVerseIndex = 0; baseVerseIndex < baseVerses.Count; baseVerseIndex++)
                {
                    var baseVerse = baseVerses[baseVerseIndex];
                    var nextBaseVerse = baseVerseIndex < baseVerses.Count - 1 ? baseVerses[baseVerseIndex + 1] : null;                   
                    
                    var getAllVerses = nextBaseVerse == null ? GetAllVersesType.All
                                                             : (nextBaseVerse.IsApocrypha != baseVerse.IsApocrypha
                                                                        ? GetAllVersesType.AllOfTheSameType
                                                                        : GetAllVersesType.One);          

                    prevParallelVerses = GetParallelVersesList(baseVerse, parallelVerses, ref parallelVerseIndex, getAllVerses, isPartVersePointer, prevParallelVerses);

                    ComparisonVersesInfo parallelVersesInfo = new ComparisonVersesInfo(prevParallelVerses);
                    //parallelVersesInfo.Strict = parallelVerses.Strict;

                    bookVersePointersComparisonTables.Add(baseVerse, parallelVersesInfo);
                }            
            }
        }

        private enum GetAllVersesType        
        {
            One,
            AllOfTheSameType,
            All
        }
        
        private static List<SimpleVersePointer> GetParallelVersesList(SimpleVersePointer baseVerse, ComparisonVersesInfo parallelVerses, ref int startIndex,
            GetAllVersesType getAllVerses, bool isPartParallelVersePointer, List<SimpleVersePointer> prevParallelVerses)
        {
            var result = new List<SimpleVersePointer>();

            var lastIndex = startIndex;
            SimpleVersePointer lastPrevVerse = prevParallelVerses.Count > 0 ? prevParallelVerses.Last() : null;
            var partIndex = lastPrevVerse != null ? lastPrevVerse.PartIndex.GetValueOrDefault(-1) + 1 : 0;

            bool getAllFirstOtherTypeVerses = startIndex == 0 && parallelVerses.Count > 0 && !baseVerse.IsApocrypha && parallelVerses[0].IsApocrypha;            

            for (int i = startIndex; i < parallelVerses.Count; i++)
            {
                if (parallelVerses[i].IsApocrypha == baseVerse.IsApocrypha || getAllVerses == GetAllVersesType.All || getAllFirstOtherTypeVerses)
                {
                    var parallelVerseToAdd = (SimpleVersePointer)parallelVerses[i].Clone();

                    if (!parallelVerseToAdd.IsApocrypha)
                        parallelVerseToAdd.PartIndex = (isPartParallelVersePointer && getAllVerses != GetAllVersesType.One) ? (int?)partIndex++ : null;
                    else if (parallelVerseToAdd.PartIndex.HasValue)
                        throw new NotSupportedException(string.Format("Apocrypha part verses are not supported yet. Parallel verse is '{0}'.", parallelVerseToAdd));                    

                    result.Add(parallelVerseToAdd);

                    lastIndex = i + 1;
                }
                else
                {
                    if (getAllVerses != GetAllVersesType.All && result.Count > 0) // то есть IsApocrypha сменилась
                        break;

                    if (getAllVerses == GetAllVersesType.AllOfTheSameType && startIndex == 0)  // то есть сразу не то пошло
                        break;
                }

                if (getAllVerses == GetAllVersesType.One && (!getAllFirstOtherTypeVerses || result.Any(v => v.IsApocrypha == baseVerse.IsApocrypha)))
                    break;                
            }
            startIndex = lastIndex;

            if (result.Count == 0)
            {
                SimpleVersePointer parallelVerseToAdd = null;

                if (!baseVerse.IsApocrypha)
                {
                    if (lastPrevVerse != null)
                    {
                        if (!lastPrevVerse.PartIndex.HasValue)
                            lastPrevVerse.PartIndex = partIndex++;

                        parallelVerseToAdd = (SimpleVersePointer)lastPrevVerse.Clone();
                        parallelVerseToAdd.PartIndex = isPartParallelVersePointer ? (int?)partIndex++ : null;
                    }
                    else
                        throw new NotSupportedException(string.Format("Can not find parallel value verse for base verse '{0}'.", baseVerse));
                }
                else
                {
                    parallelVerseToAdd = (SimpleVersePointer)baseVerse.Clone();
                    parallelVerseToAdd.IsEmpty = true;
                }

                result.Add(parallelVerseToAdd);
            }

            return result;
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
                    if (baseBookVerses == null || !baseBookVerses.ContainsKey(parallelVerseKey))   // вариант, когда и там, и там есть мы уже разобрали                    
                        bookVersePointersComparisonTables.Add(parallelVerseKey, parallelBookVerses[parallelVerseKey]);                    
                }
            }
        }
    }
}
