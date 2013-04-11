using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BibleCommon.Common
{
    public enum BibleHierarchySearchResultType
    {
        NotFound,
        PartlyFound,  // например надо было найти стих, а нашли только страницу (если искали Быт 1:120)
        Successfully,
    }

    public enum BibleHierarchyStage
    {
        SectionGroup,
        Section,
        Page,
        ContentPlaceholder
    }

    public class VerseObjectInfo
    {
        public string ObjectId { get; set; }
        public VerseNumber? VerseNumber { get; set; } // Мы, например, искали Быт 4:4 (модуль IBS). А нам вернули Быт 4:3. Здесь будем хранить "3-4".

        /// <summary>
        /// Важно! Если нет кэша Библии, то это свойство пустое
        /// </summary>
        public string ObjectHref { get; set; }

        public VerseObjectInfo()
        {
        }

        public VerseObjectInfo(VersePointerLink link)
        {
            this.ObjectId = link.ObjectId;
            this.VerseNumber = link.VerseNumber;
            this.ObjectHref = link.GetFullHref();
        }

        public bool IsVerse { get { return VerseNumber != null; } }
    }

    [Serializable]
    public class BibleHierarchyObjectInfo
    {
        public string SectionId { get; set; }
        public string PageId { get; set; }
        public string PageName { get; set; }
        public VerseObjectInfo VerseInfo { get; set; }
        public Dictionary<VersePointer, VerseObjectInfo> AdditionalObjectsIds { get; set; }  // пока заполняется только при поиске в Библии в OneNote (в HierarchySearchManager)
        public bool LoadedFromCache { get; set; }

        public List<VerseObjectInfo> GetAllObjectsIds()
        {
            var result = new List<VerseObjectInfo>();

            if (VerseInfo != null)
                result.Add(VerseInfo);

            result.AddRange(AdditionalObjectsIds.Values);

            return result;
        }

        private VerseNumber? _verseNumber;
        public VerseNumber? VerseNumber
        {
            get
            {
                if (_verseNumber.HasValue)
                    return _verseNumber;

                if (VerseInfo != null)
                    return VerseInfo.VerseNumber;

                return null;
            }
            set
            {
                _verseNumber = value;
            }
        }

        public string VerseContentObjectId
        {
            get
            {
                if (VerseInfo != null)
                    return VerseInfo.ObjectId;

                return null;
            }
        }

        public BibleHierarchyObjectInfo()
        {
            this.AdditionalObjectsIds = new Dictionary<VersePointer, VerseObjectInfo>();
        }
    }

    public class BibleSearchResult
    {
        public BibleHierarchyObjectInfo HierarchyObjectInfo { get; set; } // дополнительная информация о найденном объекте            
        public BibleHierarchyStage HierarchyStage { get; set; }
        public BibleHierarchySearchResultType ResultType { get; set; }

        public bool FoundSuccessfully
        {
            get
            {
                return ResultType == BibleHierarchySearchResultType.Successfully
                       && (HierarchyStage == BibleHierarchyStage.Page || HierarchyStage == BibleHierarchyStage.ContentPlaceholder);
            }
        }
    }
}
