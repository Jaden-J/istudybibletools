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

    public class LinkId
    {
        public string NotebookName { get; set; }

        public string PageId { get; set; }
        public string ObjectId { get; set; }

        public LinkId(string pageId, string objectId)
        {
            this.PageId = pageId;
            this.ObjectId = objectId;
        }

        public override int GetHashCode()
        {
            int result = this.PageId.GetHashCode();

            if (!string.IsNullOrEmpty(this.ObjectId))
                result = result ^ this.ObjectId.GetHashCode();

            return result;
        }

        public override bool Equals(object obj)
        {
            LinkId otherObj = (LinkId)obj;

            return this.PageId == otherObj.PageId
                && this.ObjectId == otherObj.ObjectId;
        }
    }

    public class LinkProxyInfo
    {
        public bool UseProxyLinkIfAvailable { get; set; }

        /// <summary>
        /// Использовать старый механизм NavigateToHandler (false), либо новый OneNoteProxyLinksHandler (true)
        /// </summary>
        public bool UseAdvancedProxy { get; set; }

        public bool AutoCommitLinkPage { get; set; }

        public LinkProxyInfo(bool useProxyLinkIfAvailable, bool useAdvancedProxy)
        {
            UseProxyLinkIfAvailable = useProxyLinkIfAvailable;
            UseAdvancedProxy = useAdvancedProxy;
        }
    }
}
