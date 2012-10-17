using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.Serialization;

namespace BibleCommon.Common
{
    [Serializable]
    public struct DictionaryTermLink
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

    [Serializable]
    public class VersePointerLink
    {
        public string SectionId { get; set; }
        public string PageId { get; set; }
        public string ObjectId { get; set; }
        public string Href { get; set; }
        public VerseNumber? VerseNumber { get; set; }     // Мы, например, искали Быт 4:4 (модуль IBS). А нам вернули Быт 4:3. Здесь будем хранить "3-4".
        public bool IsChapter { get; set; }
    }

    [Serializable]
    public class VersePointersCachedLinks : Dictionary<SimpleVersePointer, VersePointerLink>
    {
        public VersePointersCachedLinks()
        {
        }

        public VersePointersCachedLinks(SerializationInfo info, StreamingContext context)
            : base(info, context)
        {
        }
    }
}
