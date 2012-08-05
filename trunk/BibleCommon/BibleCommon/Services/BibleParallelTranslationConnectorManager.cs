using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BibleCommon.Common;

namespace BibleCommon.Services
{
    public static class BibleParallelTranslationConnectorManager
    {
        public static Dictionary<SimpleVersePointer, SimpleVersePointer> ConnectBibleBookTranslations(IEnumerable<BibleBookDifferences> baseBookTranslationDifferences,
            IEnumerable<BibleBookDifferences> parallelBookTranslationDifferences)
        {
            Dictionary<SimpleVersePointer, SimpleVersePointer> result = new Dictionary<SimpleVersePointer, SimpleVersePointer>();

           

            return result;
        }
    }
}
