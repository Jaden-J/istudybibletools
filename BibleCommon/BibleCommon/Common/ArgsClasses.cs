using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BibleCommon.Common
{
    public class PageIdInfo
    {
        public string SectionGroupName { get; set; }
        public string SectionName { get; set; }
        public string PageName { get; set; }
        public string PageId { get; set; }
        public string PageTitleId { get; set; }
    }
    
    public struct GetAllIncludedVersesExceptFirstArgs
    {
        public string BibleNotebookId { get; set; }
        public bool Force { get; set; }
        public bool SearchOnlyForFirstChapter { get; set; }           
    }
}
