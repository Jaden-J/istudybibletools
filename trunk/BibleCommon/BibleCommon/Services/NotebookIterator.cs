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
        public class PageInfo
        {
            public string SectionGroupId { get; set; }
            public string SectionId { get; set; }
            public string PageId { get; set; }
            public string PageName { get; set;}
        }

        private Application _oneNoteApp;

        public NotebookIterator(Application oneNoteApp)
        {
            _oneNoteApp = oneNoteApp;         
        }

        public void Iterate(string iterationProcessName, string notebookId, string sectionGroupId, Action<PageInfo> pageAction)
        {
            if (pageAction == null)
                throw new ArgumentNullException("pageAction");

            try
            {
                BibleCommon.Services.Logger.Init(iterationProcessName);

                ProcessNotebook(notebookId, sectionGroupId, pageAction);
            }
            finally
            {
                BibleCommon.Services.Logger.Done();             
            }
        }

        private void ProcessNotebook(string notebookId, string sectionGroupId, Action<PageInfo> pageAction)
        {
            BibleCommon.Services.Logger.LogMessage("Обработка записной книжки: '{0}'", 
                OneNoteUtils.GetHierarchyElementName(_oneNoteApp, notebookId));  

            string hierarchyXml;
            _oneNoteApp.GetHierarchy(notebookId, HierarchyScope.hsPages, out hierarchyXml);
            XmlNamespaceManager xnm;
            XDocument notebookDoc = OneNoteUtils.GetXDocument(hierarchyXml, out xnm);

            BibleCommon.Services.Logger.MoveLevel(1);
            ProcessRootSectionGroup(notebookId, notebookDoc, sectionGroupId, xnm, pageAction);
            BibleCommon.Services.Logger.MoveLevel(-1);
        }

        private void ProcessRootSectionGroup(string notebookId, XDocument doc, string sectionGroupId, 
            XmlNamespaceManager xnm, Action<PageInfo> pageAction)
        {
            XElement sectionGroup = string.IsNullOrEmpty(sectionGroupId)
                                        ? doc.Root
                                        : doc.Root.XPathSelectElement(
                                                string.Format("one:SectionGroup[@ID='{0}']", sectionGroupId), xnm);

            if (sectionGroup != null)
                ProcessSectionGroup(sectionGroup, sectionGroupId, notebookId, xnm, true, pageAction);
            else
                BibleCommon.Services.Logger.LogError("Не удаётся найти группу секций '{0}'", sectionGroupId);
        }

        private void ProcessSectionGroup(XElement sectionGroup, string sectionGroupId,
            string notebookId, XmlNamespaceManager xnm, bool isRootSectionGroup, Action<PageInfo> pageAction)
        {
            string sectionGroupName = (string)sectionGroup.Attribute("name");

            if (!string.IsNullOrEmpty(sectionGroupName))
            {
                BibleCommon.Services.Logger.LogMessage("Обработка группы секций '{0}'", sectionGroupName);
                BibleCommon.Services.Logger.MoveLevel(1);
            }

            if (isRootSectionGroup)
            {
                foreach (var subSectionGroup in sectionGroup.XPathSelectElements("one:SectionGroup", xnm).Where(sg => !OneNoteUtils.IsRecycleBin(sg)))
                {
                    string subSectionGroupName = (string)subSectionGroup.Attribute("name");
                    ProcessSectionGroup(subSectionGroup, subSectionGroupName, notebookId, xnm, false, pageAction);
                }
            }
            else
            {
                foreach (var subSection in sectionGroup.XPathSelectElements("one:Section", xnm))
                {
                    ProcessSection(subSection, sectionGroupId, notebookId, xnm, pageAction);
                }
            }

            if (!string.IsNullOrEmpty(sectionGroupName))
            {
                BibleCommon.Services.Logger.MoveLevel(-1);
            }
        }

        private void ProcessSection(XElement section, string sectionGroupId,
           string notebookId, XmlNamespaceManager xnm, Action<PageInfo> pageAction)
        {
            string sectionId = (string)section.Attribute("ID");
            string sectionName = (string)section.Attribute("name");

            BibleCommon.Services.Logger.LogMessage("Обработка секции '{0}'", sectionName);
            BibleCommon.Services.Logger.MoveLevel(1);

            foreach (var page in section.XPathSelectElements("one:Page", xnm))
            {
                string pageId = (string)page.Attribute("ID");
                string pageName = (string)page.Attribute("name");

                BibleCommon.Services.Logger.LogMessage("Обработка страницы '{0}'", pageName);

                BibleCommon.Services.Logger.MoveLevel(1);

                pageAction(new PageInfo()
                            {
                                SectionGroupId = sectionGroupId,
                                SectionId = sectionId,
                                PageId = pageId,
                                PageName = pageName
                            });

                BibleCommon.Services.Logger.MoveLevel(-1);                
            }

            BibleCommon.Services.Logger.MoveLevel(-1);
        }       
    }
}
