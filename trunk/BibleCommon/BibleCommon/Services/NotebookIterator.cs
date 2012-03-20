using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.OneNote;
using BibleCommon.Helpers;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;

namespace BibleCommon.Services
{
    public class NotebookIterator
    {
        public class HierarchyElementInfo
        {
            public string Id { get; set; }
            public string Title { get; set; }
        }

        public class SectionGroupInfo : HierarchyElementInfo
        {            
            public List<SectionGroupInfo> SectionGroups { get; set; }
            public List<SectionInfo> Sections { get; set; }

            public SectionGroupInfo()
            {
                this.SectionGroups = new List<SectionGroupInfo>();
                this.Sections = new List<SectionInfo>();
            }
        }
   

        public class NotebookInfo : SectionGroupInfo
        {  
            public int PagesCount { get; set; }

            public NotebookInfo(SectionGroupInfo container)
            {
                this.SectionGroups = container.SectionGroups;
                this.Sections = container.Sections;
                this.Title = container.Title;                
            }
        }

        public class SectionInfo : HierarchyElementInfo
        {
            public List<PageInfo> Pages { get; set; }

            public SectionInfo()                
            {
                this.Pages = new List<PageInfo>();
            }
        }

        public class PageInfo : HierarchyElementInfo
        {
            public string SectionGroupId { get; set; }
            public string SectionId { get; set; }            
        }

        private Application _oneNoteApp;

        public NotebookIterator(Application oneNoteApp)
        {
            _oneNoteApp = oneNoteApp;         
        }

        public NotebookInfo GetNotebookPages(string notebookId, string sectionGroupId, Func<PageInfo, bool> filter)
        {
            OneNoteProxy.HierarchyElement notebookElement = OneNoteProxy.Instance.GetHierarchy(_oneNoteApp, notebookId, HierarchyScope.hsPages);

            XElement sectionGroup = string.IsNullOrEmpty(sectionGroupId)
                                        ? notebookElement.Content.Root
                                        : notebookElement.Content.Root.XPathSelectElement(
                                                string.Format("one:SectionGroup[@ID='{0}']", sectionGroupId), notebookElement.Xnm);

            if (sectionGroup == null)
                throw new Exception(string.Format("Не удаётся найти группу секций '{0}'", sectionGroupId));
            
            int pagesCount = 0;
            var result = new NotebookInfo(ProcessSectionGroup(sectionGroup, notebookId, notebookElement.Xnm, ref pagesCount));
            result.Id = notebookId;
            result.PagesCount = pagesCount;
            return result;
        }

        public void Iterate(string notebookId, string sectionGroupId, Action<PageInfo> pageAction)
        {
            if (pageAction == null)
                throw new ArgumentNullException("pageAction");

            BibleCommon.Services.Logger.LogMessage("Обработка записной книжки: '{0}'",
                OneNoteUtils.GetHierarchyElementName(_oneNoteApp, notebookId));
            
            NotebookInfo notebook = GetNotebookPages(notebookId, sectionGroupId, null);

            BibleCommon.Services.Logger.MoveLevel(1);
            IterateContainer(notebook, pageAction);
            BibleCommon.Services.Logger.MoveLevel(-1);
        }

        private void IterateContainer(SectionGroupInfo container, Action<PageInfo> pageAction)
        {
            if (container is SectionGroupInfo)
            {
                BibleCommon.Services.Logger.LogMessage("Обработка группы секций '{0}'", container.Title);
                BibleCommon.Services.Logger.MoveLevel(1);
            }

            foreach (SectionInfo section in container.Sections)
            {
                BibleCommon.Services.Logger.LogMessage("Обработка секции '{0}'", section.Title);
                BibleCommon.Services.Logger.MoveLevel(1);

                foreach (PageInfo page in section.Pages)
                {
                    BibleCommon.Services.Logger.LogMessage("Обработка страницы '{0}'", page.Title);
                    BibleCommon.Services.Logger.MoveLevel(1);

                    pageAction(page);

                    BibleCommon.Services.Logger.MoveLevel(-1);
                }

                BibleCommon.Services.Logger.MoveLevel(-1);
            }

            foreach (SectionGroupInfo sectionGroup in container.SectionGroups)
            {
                IterateContainer(sectionGroup, pageAction);
            }

            if (container is SectionGroupInfo)
                BibleCommon.Services.Logger.MoveLevel(-1);
        }


        private SectionGroupInfo ProcessSectionGroup(XElement sectionGroupElement, string notebookId, XmlNamespaceManager xnm, ref int pagesCount)
        {
            SectionGroupInfo sectionGroup = new SectionGroupInfo();
            ProcessHierarchyElement(sectionGroup, sectionGroupElement);            

            foreach (var subSectionGroup in sectionGroupElement.XPathSelectElements("one:SectionGroup", xnm)
                .Where(sg => !OneNoteUtils.IsRecycleBin(sg)))
            {                
                sectionGroup.SectionGroups.Add(ProcessSectionGroup(subSectionGroup, notebookId, xnm, ref pagesCount));
            }

            foreach (var subSection in sectionGroupElement.XPathSelectElements("one:Section", xnm))
            {
                sectionGroup.Sections.Add(ProcessSection(subSection, sectionGroup.Id, notebookId, xnm, ref pagesCount));
            }

            return sectionGroup;
        }

        private SectionInfo ProcessSection(XElement sectionElement, string sectionGroupId,
           string notebookId, XmlNamespaceManager xnm, ref int pagesCount)
        {
            SectionInfo section = new SectionInfo();
            ProcessHierarchyElement(section, sectionElement);            

            foreach (var pageElement in sectionElement.XPathSelectElements("one:Page", xnm))
            {
                if (!OneNoteUtils.IsRecycleBin(pageElement))
                {
                    var page = new PageInfo()
                                {
                                    SectionGroupId = sectionGroupId,
                                    SectionId = section.Id,                                 
                                };                    
                    ProcessHierarchyElement(page, pageElement);

                    section.Pages.Add(page);
                    pagesCount++;
                }
            }

            return section;
        }

        public void ProcessHierarchyElement(HierarchyElementInfo hierarchyElement, XElement xElement)
        {
            hierarchyElement.Title = (string)xElement.Attribute("name");
            hierarchyElement.Id = (string)xElement.Attribute("ID");
        }

    }
}
