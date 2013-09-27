using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace BibleCommon.Common
{  
    
    public struct GetAllIncludedVersesExceptFirstArgs
    {
        public string BibleNotebookId { get; set; }
        public bool Force { get; set; }
        public bool SearchOnlyForFirstChapter { get; set; }

        public bool TryToGroupVersesInChapters { get; set; }   // если ссылка "Быт 1:5-4:3", то вернёт Быт 2 и Быт 3 как главы без стихов 
    }
   
}
