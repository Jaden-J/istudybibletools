using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BibleCommon.Helpers;

namespace BibleCommon.Common
{
    public class BibleTranslationDifferencesFormulaBase
    {
        public string OriginalFormula { get; set; }

        public BibleTranslationDifferencesFormulaBase(string originalFormula)
        {
            if (string.IsNullOrEmpty(originalFormula))
                throw new ArgumentNullException("Formula is null");

            this.OriginalFormula = originalFormula;

            int indexOfColon = originalFormula.IndexOf(":");

            if (indexOfColon == -1)
                throw new NotSupportedException(string.Format("Unknown formula: '{0}'", originalFormula));

            if (indexOfColon != originalFormula.LastIndexOf(":"))
                throw new NotSupportedException(
                    string.Format("The verses formula with two colons (':') is not supported yet: '{0}'", originalFormula));
        }
    }

    public class BibleTranslationDifferencesBaseVersesFormula : BibleTranslationDifferencesFormulaBase
    {
        public int BookIndex { get; set; }
        protected VersePointer BaseVersePointer { get; set; }

        internal bool IsMultiVerse
        {
            get
            {
                return BaseVersePointer.IsMultiVerse;
            }
        }

        internal int FirstVerse
        {
            get
            {
                return BaseVersePointer.Verse.GetValueOrDefault(0);
            }
        }

        internal int LastVerse
        {
            get
            {
                return BaseVersePointer.TopVerse.GetValueOrDefault(0);
            }
        }

        internal int FirstChapter
        {
            get
            {
                return BaseVersePointer.Chapter.GetValueOrDefault(0);
            }
        }

        internal int LastChapter
        {
            get
            {
                return BaseVersePointer.TopChapter.GetValueOrDefault(0);
            }
        }
            

        public BibleTranslationDifferencesBaseVersesFormula(int bookIndex, string baseVersesFormula)
            : base(baseVersesFormula)
        {
            this.BookIndex = bookIndex;
            Initialize(baseVersesFormula);
        }     

        private List<SimpleVersePointer> _allVerses;
        public List<SimpleVersePointer> GetAllVerses()
        {
            if (_allVerses == null)
            {
                _allVerses = new List<SimpleVersePointer>();
                _allVerses.Add(new SimpleVersePointer(BookIndex, FirstChapter, FirstVerse));

                if (IsMultiVerse)
                    _allVerses.AddRange(BaseVersePointer.GetAllIncludedVersesExceptFirst(null, null, true)
                        .ConvertAll<SimpleVersePointer>(v => new SimpleVersePointer(BookIndex, v.Chapter.GetValueOrDefault(), v.Verse.GetValueOrDefault(0))));
            }

            return _allVerses;
        }

        ///// <summary>
        ///// Сделать из MultiVerse обычный, взяв только первый стих
        ///// </summary>
        //public void TrimVerses()
        //{
        //    Initialize(string.Format("{0}:{1}", FirstChapter, FirstVerse));
        //}

        private void Initialize(string baseVersesFormula)
        {
            BaseVersePointer = new VersePointer("book " + baseVersesFormula);
            _allVerses = null;
        }
    }    

    public class BibleTranslationDifferencesParallelVersesFormula : BibleTranslationDifferencesFormulaBase
    {
        public abstract class ParallelFormulaPart
        {
            public int? Deviation { get; set; }
            public List<int> ParallelVerses { get; set; }
            public bool? IsPartVersePointer { get; set; }            

            public string FormulaPart { get; set; }
            public BibleTranslationDifferencesParallelVersesFormula ParallelVersesFormula { get; set; }               

            protected int? FirstVerse { get; set; }
            protected int? LastVerse { get; set; }

            public ParallelFormulaPart(string formulaPart, BibleTranslationDifferencesParallelVersesFormula parallelVersesFormula)
            {
                this.FormulaPart = formulaPart;
                this.ParallelVersesFormula = parallelVersesFormula;                

                int indexOfX = formulaPart.IndexOf("X");

                if (indexOfX == -1)
                {
                    int indexOfDash = formulaPart.IndexOf("-");
                    if (indexOfDash == -1)   // тогда просто число
                    {
                        ParseSimpleVerse();                        
                    }
                    else  // тогда диапазон стихов, типа "1-3"
                    {
                        ParseMultiVerse();                       
                    }
                }
                else  // тогда формула, типа "X+1"
                {
                    ParseFormulaVerse();
                }
            }

            protected virtual void ParseFormulaVerse()
            {
                int? value = StringUtils.GetStringLastNumber(FormulaPart);

                if (!value.HasValue)
                {
                    if (FormulaPart.Length == 1)  // только "X"
                        Deviation = 0;
                    else                    
                        ThrowUnsupportedFormulaException("x004");                    
                }
                else
                {
                    Deviation = value.Value;
                }

                if (FormulaPart.IndexOf("+") == -1)
                    Deviation *= -1;
            }

            protected virtual void ParseMultiVerse()
            {
                FirstVerse = StringUtils.GetStringFirstNumber(FormulaPart);
                LastVerse = StringUtils.GetStringLastNumber(FormulaPart);

                if (!FirstVerse.HasValue || !LastVerse.HasValue || FirstVerse.Value == LastVerse.Value)
                    ThrowUnsupportedFormulaException("x002");
            }

            protected virtual void CalculateDeviationForMultiVerseFormula(int firstBaseVerse, int lastBaseVerse)
            {
                Deviation = FirstVerse - firstBaseVerse;

                if (Deviation != (LastVerse - lastBaseVerse))
                    ThrowUnsupportedFormulaException("x003");
            }

            protected virtual void ParseSimpleVerse()
            {
                int value;
                if (!int.TryParse(FormulaPart, out value))
                    ThrowUnsupportedFormulaException("x001");

                ParallelVerses = new List<int>() { int.Parse(FormulaPart) };                
            }

            protected virtual void ThrowUnsupportedFormulaException(string errorCode)
            {
                throw new NotSupportedException(string.Format("Not supported parallel verses formula part (errorCode = {0}): '{1}'", errorCode, FormulaPart));
            }           
        }

        public class ParallelChapterFormulaPart : ParallelFormulaPart
        {
            public ParallelChapterFormulaPart(string formulaPart, BibleTranslationDifferencesParallelVersesFormula parallelVersesFormula)
                : base(formulaPart, parallelVersesFormula)
            {
            }

            protected override void ParseMultiVerse()
            {
                base.ParseMultiVerse();
                
                int firstBaseVerse = ParallelVersesFormula.BaseVersesFormula.FirstChapter;
                int lastBaseVerse = ParallelVersesFormula.BaseVersesFormula.LastChapter;

                CalculateDeviationForMultiVerseFormula(firstBaseVerse, lastBaseVerse);
            }

            public int CalculateParallelChapter(int baseChapter)
            {
                if (Deviation.HasValue)
                    return baseChapter + Deviation.Value;
                else
                {
                    if (ParallelVerses.Count != 1)
                        ThrowUnsupportedFormulaException("x005");

                    return ParallelVerses[0];
                }
            }       
        }

        public class ParallelVersesFormulaPart : ParallelFormulaPart
        {
            protected BibleBookDifference.VerseAlign VerseAlign { get; set; }
            protected bool StrictProcessing { get; set; }

            public ParallelVersesFormulaPart(string formulaPart,
                BibleTranslationDifferencesParallelVersesFormula parallelVersesFormula, BibleBookDifference.VerseAlign verseAlign, bool strictProcessing)
                : base(formulaPart, parallelVersesFormula)
            {
                this.VerseAlign = verseAlign;
                this.StrictProcessing = strictProcessing;
            }

            protected override void ParseMultiVerse()
            {
                base.ParseMultiVerse();

                if (!ParallelVersesFormula.BaseVersesFormula.IsMultiVerse)         // например в случае "1:1 -> 1:2-3" 
                {
                    ParallelVerses = new List<int>();
                    for (int i = FirstVerse.Value; i <= LastVerse.Value; i++)
                        ParallelVerses.Add(i);
                }
                else
                {
                    int firstBaseVerse = ParallelVersesFormula.BaseVersesFormula.FirstVerse;
                    int lastBaseVerse = ParallelVersesFormula.BaseVersesFormula.LastVerse;

                    CalculateDeviationForMultiVerseFormula(firstBaseVerse, lastBaseVerse);
                }
            }

            protected override void ParseSimpleVerse()
            {
                base.ParseSimpleVerse();

                if (ParallelVersesFormula.BaseVersesFormula.IsMultiVerse)         // например в случае "1:1-2 -> 1:1"                        
                {
                    IsPartVersePointer = true;
                }
            }

            public List<int> CalculateParallelVerses(int baseVerse)
            {
                if (Deviation.HasValue)
                    return new List<int>() { baseVerse + Deviation.Value };
                else
                    return ParallelVerses;
            }
        }

        protected BibleTranslationDifferencesBaseVersesFormula BaseVersesFormula { get; set; }
        protected ParallelChapterFormulaPart ChapterFormulaPart { get; set; }
        protected ParallelVersesFormulaPart VersesFormulaPart { get; set; }

        public BibleTranslationDifferencesParallelVersesFormula(string parallelVersesFormula,
            BibleTranslationDifferencesBaseVersesFormula baseVersesFormula, BibleBookDifference.VerseAlign verseAlign, bool strictProcessing)
            : base(parallelVersesFormula)
        {
            this.BaseVersesFormula = baseVersesFormula;

            int indexOfColon = parallelVersesFormula.IndexOf(":");
            ChapterFormulaPart = new ParallelChapterFormulaPart(parallelVersesFormula.Substring(0, indexOfColon), this);
            VersesFormulaPart = new ParallelVersesFormulaPart(parallelVersesFormula.Substring(indexOfColon + 1), this, verseAlign, strictProcessing);
        }

        public ComparisonVersesInfo GetParallelVerses(SimpleVersePointer baseVerse, SimpleVersePointer prevVerse)
        {
            var result = new ComparisonVersesInfo();

            if (ChapterFormulaPart != null && VersesFormulaPart != null)
            {
                int versePartIndex = prevVerse != null ? prevVerse.PartIndex.GetValueOrDefault(0) + 1 : 0;
                foreach (var parallelVerse in VersesFormulaPart.CalculateParallelVerses(baseVerse.Verse))                    
                {
                    var parallelVersePointer = new SimpleVersePointer(
                        baseVerse.BookIndex, ChapterFormulaPart.CalculateParallelChapter(baseVerse.Chapter), parallelVerse);

                    if (VersesFormulaPart.IsPartVersePointer.GetValueOrDefault(false))
                    {                        
                        parallelVersePointer.PartIndex = versePartIndex++;
                    }

                    result.Add(parallelVersePointer);
                }
            }

            return result;
        }
    }
}
