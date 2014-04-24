using BibleCommon.Consts;
using BibleCommon.Handlers;
using BibleCommon.Helpers;
using BibleCommon.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.XPath;

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
        public string NotebookName { get; set; }

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

        public int GetLevel()
        {
            if (Parent != null)
                return Parent.GetLevel() + 1;
            else
                return 1;
        }

        public void LoadHierarchyElementParent(string notebookId, ApplicationCache.HierarchyElement fullNotebookHierarchy)
        {
            var el = fullNotebookHierarchy.Content.Root.XPathSelectElement(
                                string.Format("//one:{0}[@ID=\"{1}\"]", this.GetElementName(), this.Id), fullNotebookHierarchy.Xnm);

            if (el == null)
                throw new Exception(string.Format("Can not find hierarchyElement '{0}' of type '{1}' in notebook '{2}'",
                                this.Id, this.Type, notebookId));


            if (el.Parent != null)
            {
                var parentId = (string)el.Parent.Attribute("ID");
                var parentType = (HierarchyElementType)Enum.Parse(typeof(HierarchyElementType), el.Parent.Name.LocalName);

                string parentName;
                string parentTitle;
                if (parentType == HierarchyElementType.Notebook)
                {
                    parentTitle = (string)el.Parent.Attribute("nickname");
                    parentName = (string)el.Parent.Attribute("name");

                    if (string.IsNullOrEmpty(parentTitle))
                        parentTitle = parentName;
                }
                else
                {
                    parentName = (string)el.Parent.Attribute("name");
                    parentTitle = parentName;
                }

                var parent = new HierarchyElementInfo() { Id = parentId, Title = parentTitle, Name = parentName, Type = parentType, NotebookId = notebookId };
                parent.LoadHierarchyElementParent(notebookId, fullNotebookHierarchy);
                this.Parent = parent;
            }
        }       
    }

    public class PageHierarchyInfo : HierarchyElementInfo
    {
        public void LoadPageSyncId(ApplicationCache.PageContent notePageDocument)
        {
            this.SyncPageId = OneNoteUtils.GetElementMetaData(notePageDocument.Content.Root, Constants.Key_SyncId, notePageDocument.Xnm);
            if (string.IsNullOrEmpty(this.SyncPageId))
                this.SyncPageId = OneNoteUtils.GetElementMetaData(notePageDocument.Content.Root, Constants.Key_PId, notePageDocument.Xnm);

            if (string.IsNullOrEmpty(this.SyncPageId))
            {
                this.SyncPageId = OneNoteProxyLinksHandler.GeneratePId();
                OneNoteUtils.UpdateElementMetaData(notePageDocument.Content.Root, Constants.Key_PId, this.SyncPageId, notePageDocument.Xnm);
            }
        }       
    }
}
