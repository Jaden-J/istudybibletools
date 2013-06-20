using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BibleCommon.Common;

namespace BibleCommon.Services
{
    public class AnalyzedVersesService
    {
        public AnalyzedVersesInfo VersesInfo { get; set; }

        public AnalyzedVersesService()
        {

        }

        public void UpdateVerseInfo(SimpleVersePointer verse, decimal weight)
        {
            VersesInfo
                .GetOrCreateBookInfo(verse.BookIndex)
                    .GetOrCreateChapterInfo(verse.Chapter)
        }

        public void Update()
        {

        }
    }
}
