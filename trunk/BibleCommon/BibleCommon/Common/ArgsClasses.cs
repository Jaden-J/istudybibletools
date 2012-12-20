using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BibleCommon.Common
{
    public enum HierarchyElementType
    {
        Notebook,
        SectionGroup,
        Section,
        Page
    }

    public class HierarchyElementInfo
    {   
        public string Name { get; set; }
        public string Id { get; set; }
        public HierarchyElementType Type { get; set; }
        
        public HierarchyElementInfo Parent { get; set; }
        public string PageTitleId { get; set; }
        public string NotebookId { get; set; }

        public string GetElementName()
        {
            return GetElementName(this.Type);
        }

        public static string GetElementName(HierarchyElementType type)
        {
            switch (type)
            {
                case HierarchyElementType.Notebook:
                    return "Notebook";
                case HierarchyElementType.SectionGroup:
                    return "SectionGroup";
                case HierarchyElementType.Section:
                    return "Section";
                case HierarchyElementType.Page:
                    return "Page";
                default:
                    throw new NotSupportedException(type.ToString());
            }
        }

        [Obsolete]
        public string SectionName
        {
            get
            {
                if (Type == HierarchyElementType.Page && Parent != null)
                    return Parent.Name;

                return null;
            }
        }

        [Obsolete]
        public string SectionGroupName
        {
            get
            {
                if (Type == HierarchyElementType.Page && Parent != null && Parent.Parent != null)
                    return Parent.Parent.Name;

                return null;
            }
        }        
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
