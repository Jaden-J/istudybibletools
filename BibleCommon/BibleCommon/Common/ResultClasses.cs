using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BibleCommon.Common
{
    internal class ProcessFoundVerseResult
    {
        internal int CursorPosition { get; set; }
        internal bool WasModified { get; set; }
    }

    public class ErrorsList : List<string>
    {
        public string ErrorsDecription { get; set; }

        public ErrorsList(IEnumerable<string> collection)
            : base(collection)
        {
        }
    }
}
