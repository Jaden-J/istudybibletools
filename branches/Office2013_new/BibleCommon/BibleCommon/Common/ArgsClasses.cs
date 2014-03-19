using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace BibleCommon.Common
{  
    
    public struct GetAllIncludedVersesArgs
    {
        public string BibleNotebookId { get; set; }
        public bool Force { get; set; }
        public bool SearchOnlyForFirstChapter { get; set; }

        public bool TryToGroupVersesInChapters { get; set; }   // если ссылка "Быт 1:5-4:3", то вернёт Быт 2 и Быт 3 как главы без стихов 
    }

    public class GetAllIncludedVersesResult
    {
        public List<VersePointer> Verses { get; set; }
        public int VersesCount { get; set; }
        public bool? NotNeedFirstVerse { get; set; }
        
        public GetAllIncludedVersesResult()
        {
            Verses = new List<VersePointer>();
        }
    }

    public struct LinkProxyInfo
    {
        public bool UseProxyLinkIfAvailable { get; set; }

        /// <summary>
        /// Использовать старый механизм NavigateToHandler (true), либо новый OneNoteProxyLinksHandler (false)
        /// </summary>
        public bool UseSimpleProxy { get; set; }
    }
}
