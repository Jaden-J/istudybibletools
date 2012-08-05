using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BibleCommon.Common
{
    public class BibleTranslationDifferencesEx
    {
        public class SimpleVersePointersComparisonTable : Dictionary<SimpleVersePointer, SimpleVersePointer>
        {
        }

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
            VersePointer vp = new VersePointer("book " + bookDifference.BaseVerses);

            List<SimpleVersePointer> verses = new List<SimpleVersePointer>();
            verses.Add(new SimpleVersePointer(bookIndex, vp.Chapter.GetValueOrDefault(0), vp.Verse.GetValueOrDefault(0)));

            if (vp.IsMultiVerse)
                verses.AddRange(vp.GetAllIncludedVersesExceptFirst(null, null, true)
                    .ConvertAll<SimpleVersePointer>(v => new SimpleVersePointer(bookIndex, v.Chapter.GetValueOrDefault(), v.Verse.GetValueOrDefault(0))));

            foreach (var verse in verses)
            {
                SimpleVersePointer parallelVerse = CalculateParallelVerse(verse, bookDifference.Verses);
                BibleVersesDifferences[bookIndex].Add(verse, parallelVerse);
            }
        }

        private SimpleVersePointer CalculateParallelVerse(SimpleVersePointer verse, string parallelVerseFormula)
        {
            if (parallelVerseFormula == string.Empty)
                return null;

            int indexOfColon = parallelVerseFormula.IndexOf(":");

            if (indexOfColon == -1)
                throw new NotSupportedException(string.Format("Unknown formula: '{0}'", parallelVerseFormula));            

            string firstPart = parallelVerseFormula.Substring(0, indexOfColon);
            string secondPart = parallelVerseFormula.Substring(indexOfColon + 1);

            int parallelChapter = CalculateFormula(verse.Chapter, firstPart);
            int parallelVerse = CalculateFormula(verse.Verse, secondPart);                        

            return new SimpleVersePointer(verse.BookIndex, parallelChapter, parallelVerse);
        }

        private int CalculateFormula(int baseNumber, string formula)
        {
            int indexOfX = formula.IndexOf("X");

            if (indexOfX == -1)
            {
                проблема в том, что могут быть и такие: 40:1-5 -> 39:31-35
            }

            int 


        }
    }
}
