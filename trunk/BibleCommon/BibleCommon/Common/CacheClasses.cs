using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.Serialization;
using BibleCommon.Services;

namespace BibleCommon.Common
{
    [Serializable]
    public class DictionaryTermLink
    {
        private string _separator = "_|_";

        public string PageId { get; set; }
        public string ObjectId { get; set; }
        public string Href { get; set; }

        public DictionaryTermLink()
        {
        }

        public DictionaryTermLink(string s)
        {
            var parts = s.Split(new string[] { _separator }, StringSplitOptions.None);

            PageId = parts[0];
            ObjectId = parts[1];
            Href = parts[2];
        }

        public override string ToString()
        {
            return string.Join(_separator, new string[] { PageId, ObjectId, Href });
        }
    }    

    [Serializable]
    public class VersePointerLink
    {
        private string _separator = "_|_";

        public string SectionId { get; set; }
        public string PageId { get; set; }
        public string PageName { get; set; }
        public string ObjectId { get; set; }
        public string Href { get; set; }
        public VerseNumber? VerseNumber { get; set; }     // Мы, например, искали Быт 4:4 (модуль IBS). А нам вернули Быт 4:3. Здесь будем хранить "3-4".
        public bool IsChapter { get; set; }
        

        public VersePointerLink()
        {
        }

        public VersePointerLink(string s)
        {
            var parts = s.Split(new string[] { _separator }, StringSplitOptions.None);

            SectionId = parts[0];
            PageId = parts[1];
            PageName = parts[2];
            ObjectId = parts[3];
            Href = parts[4];
            VerseNumber = string.IsNullOrEmpty(parts[5]) ? null : (VerseNumber?) Common.VerseNumber.Parse(parts[5]);
            IsChapter = bool.Parse(parts[6]);
        }

        public override string ToString()
        {
            return string.Join(_separator, new string[] { SectionId, PageId, PageName, ObjectId, Href, VerseNumber.HasValue ? VerseNumber.Value.ToString() : string.Empty, IsChapter.ToString() });
        }
    }   
}
