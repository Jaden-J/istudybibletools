using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.Serialization;

namespace BibleCommon.Common
{
    [Serializable]
    public class DictionaryTermLink
    {
        public string PageId { get; set; }
        public string ObjectId { get; set; }
        public string Href { get; set; }
    }

    [Serializable]
    public class DictionaryCachedTermSet : Dictionary<string, DictionaryTermLink>
    {
        public DictionaryCachedTermSet()
        {
        }

        public DictionaryCachedTermSet(SerializationInfo info, StreamingContext context)
            : base(info, context)
        {
        }
    }
}
