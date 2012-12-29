using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace BibleCommon.Common
{
    public struct XmlCursorPosition : IComparable
    {
        public int LineNumber;
        public int LinePosition;

        public int CompareTo(object obj)
        {
            var otherObj = (XmlCursorPosition)obj;

            var result = this.LineNumber.CompareTo(otherObj.LineNumber);
            if (result == 0)
                result = this.LinePosition.CompareTo(otherObj.LinePosition);

            return result;
        }

        public XmlCursorPosition(IXmlLineInfo lineInfo)
        {
            LineNumber = lineInfo.LineNumber;
            LinePosition = lineInfo.LinePosition;
        }

        public XmlCursorPosition(string lineInfo)
        {
            var parts = lineInfo.Split(new char[] { ';' });

            LineNumber = int.Parse(parts[0]);
            LinePosition = int.Parse(parts[1]);
        }

        public override bool Equals(object obj)
        {
            if (!(obj is XmlCursorPosition))
                return false;

            var otherObj = (XmlCursorPosition)obj;

            return this.LineNumber == otherObj.LineNumber
                && this.LinePosition == otherObj.LinePosition;
        }

        public override string ToString()
        {
            return string.Format("{0};{1}", this.LineNumber, this.LinePosition);
        }

        public override int GetHashCode()
        {
            return this.LineNumber.GetHashCode() ^ this.LinePosition.GetHashCode();
        }

        public static bool operator >(XmlCursorPosition cp1, XmlCursorPosition cp2)
        {
            return cp1.CompareTo(cp2) > 0;
        }

        public static bool operator <(XmlCursorPosition cp1, XmlCursorPosition cp2)
        {
            return cp1.CompareTo(cp2) < 0;
        }

        public static bool operator ==(XmlCursorPosition cp1, XmlCursorPosition cp2)
        {
            return cp1.Equals(cp2);
        }

        public static bool operator !=(XmlCursorPosition cp1, XmlCursorPosition cp2)
        {
            return !cp1.Equals(cp2);
        }        
    }


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

        private string _uniqueName;
        /// <summary>
        /// Имя, используемое на странице сводной заметок
        /// </summary>
        public string UniqueName
        {
            get
            {
                if (string.IsNullOrEmpty(_uniqueName))
                    return Name;
                else
                    return _uniqueName;
            }
            set
            {
                _uniqueName = value;
            }
        }

        
        /// <summary>
        /// Идентификатор, идентифицирующий заметку на странице сводной заметок
        /// </summary>
        public string UniqueId { get; set; }        

        public HierarchyElementType Type { get; set; }
        
        public HierarchyElementInfo Parent { get; set; }
        public string PageTitleId { get; set; }

        private string _uniqueNoteTitleId;
        public string UniqueNoteTitleId
        {
            get
            {
                if (string.IsNullOrEmpty(_uniqueNoteTitleId))
                    return PageTitleId;
                else
                    return _uniqueNoteTitleId;
            }
            set
            {
                _uniqueNoteTitleId = value;
            }
        }    
   

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
