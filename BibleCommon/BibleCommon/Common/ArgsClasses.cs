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
    }
   
}
