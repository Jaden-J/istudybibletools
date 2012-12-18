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

    public class ItemInfo
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string Nickname { get; set; }

        public ItemInfo(string id, string name)
        {
            this.Id = id;
            this.Name = name;
        }

        public override string ToString()
        {
            if (!string.IsNullOrEmpty(Nickname))
                return Nickname;
            else
                return Name;
        }

        public override int GetHashCode()
        {
            return Id.GetHashCode() ^ Name.GetHashCode();
        }

        public override bool Equals(object obj)
        {
            if (!(obj is ItemInfo))
                return false;

            return ((ItemInfo)obj).Id == this.Id;
        }
    }
}
