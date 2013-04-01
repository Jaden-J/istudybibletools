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
        /// <summary>
        /// Отображаемое имя
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// Идентификационное имя. Важно для записных книжек
        /// </summary>
        public string Name { get; set; }

        public string SyncPageId { get; set; }  // уникальный ID страницы, который синхронизируется! (хранится в метаданных). Относится только к странице.        

        public string Id { get; set; }

        public string UniqueName   //  уникальный идентификатор
        {
            get
            {
                switch (Type)
                {
                    case HierarchyElementType.Notebook:
                    case HierarchyElementType.SectionGroup:
                    case HierarchyElementType.Section:
                        return Name;   // идентифицируем по Name, так как ID на разных компьютерах разный
                    case HierarchyElementType.Page:
                        return ManualId ?? SyncPageId ?? Id;
                }

                throw new NotSupportedException(Type.ToString());
            }
        }

        private string _uniqueTitle;
        /// <summary>
        /// Имя, используемое на странице сводной заметок
        /// </summary>
        public string UniqueTitle
        {
            get
            {
                if (string.IsNullOrEmpty(_uniqueTitle))
                    return Title;
                else
                    return _uniqueTitle;
            }
            set
            {
                _uniqueTitle = value;
            }
        }


        /// <summary>
        /// Идентификатор, идентифицирующий заметку на странице сводной заметок. Если не задан - используется Id.
        /// </summary>
        public string ManualId { get; set; }

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
                    return Parent.Title;

                return null;
            }
        }

        [Obsolete]
        public string SectionGroupName
        {
            get
            {
                if (Type == HierarchyElementType.Page && Parent != null && Parent.Parent != null)
                    return Parent.Parent.Title;

                return null;
            }
        }
    }
}
