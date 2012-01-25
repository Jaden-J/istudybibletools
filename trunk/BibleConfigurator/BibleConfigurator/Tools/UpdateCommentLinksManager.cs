using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.OneNote;
using BibleCommon.Helpers;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;
using BibleCommon.Services;

namespace BibleConfigurator.Tools
{
    public class RelinkAllBibleCommentsManager
    {
        private Application _oneNoteApp;

        public RelinkAllBibleCommentsManager(Application oneNoteApp)
        {
            _oneNoteApp = oneNoteApp; 
        }

        public void RelinkAllBibleComments()
        {
            BibleCommon.Services.Logger.Init("RelinkAllBibleCommentsManager");

            ProcessNotebook(SettingsManager.Instance.NotebookId_Bible, SettingsManager.Instance.SectionGroupId_Bible);

            BibleCommon.Services.Logger.Done();
        }

        private void ProcessNotebook(string notebookId, string sectionGroupId)
        {
            BibleCommon.Services.Logger.LogMessage("Обработка записной книжки: '{0}'", OneNoteUtils.GetHierarchyElementName(_oneNoteApp, notebookId));  // чтобы точно убедиться

            string hierarchyXml;
            _oneNoteApp.GetHierarchy(notebookId, HierarchyScope.hsPages, out hierarchyXml);
            XmlNamespaceManager xnm;
            XDocument notebookDoc = OneNoteUtils.GetXDocument(hierarchyXml, out xnm);

            BibleCommon.Services.Logger.MoveLevel(1);
            ProcessRootSectionGroup(notebookId, notebookDoc, sectionGroupId, xnm);
            BibleCommon.Services.Logger.MoveLevel(-1);
        }

        private void ProcessRootSectionGroup(string notebookId, XDocument doc, string sectionGroupId, XmlNamespaceManager xnm)
        {
            XElement sectionGroup = string.IsNullOrEmpty(sectionGroupId)
                                        ? doc.Root
                                        : doc.Root.XPathSelectElement(
                                                string.Format("one:SectionGroup[@ID='{0}']", sectionGroupId), xnm);

            if (sectionGroup != null)
                ProcessSectionGroup(sectionGroup, sectionGroupId, notebookId, xnm);
            else
                BibleCommon.Services.Logger.LogError("Не удаётся найти группу секций '{0}'", sectionGroupId);
        }

        private  void ProcessSectionGroup(XElement sectionGroup, string sectionGroupId,
            string notebookId, XmlNamespaceManager xnm)
        {
            string sectionGroupName = (string)sectionGroup.Attribute("name");

            if (!string.IsNullOrEmpty(sectionGroupName))
            {
                BibleCommon.Services.Logger.LogMessage("Обработка группы секций '{0}'", sectionGroupName);
                BibleCommon.Services.Logger.MoveLevel(1);
            }

            foreach (var subSectionGroup in sectionGroup.XPathSelectElements("one:SectionGroup", xnm))
            {
                string subSectionGroupName = (string)subSectionGroup.Attribute("name");
                ProcessSectionGroup(subSectionGroup, subSectionGroupName, notebookId, xnm);
            }

            foreach (var subSection in sectionGroup.XPathSelectElements("one:Section", xnm))
            {
                ProcessSection(subSection, sectionGroupId, notebookId, xnm);
            }

            if (!string.IsNullOrEmpty(sectionGroupName))
            {
                BibleCommon.Services.Logger.MoveLevel(-1);
            }
        }

        private void ProcessSection(XElement section, string sectionGroupId,
           string notebookId, XmlNamespaceManager xnm)
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

                RelinkPageComments(sectionGroupId, sectionId, pageId, pageName);
                
                BibleCommon.Services.Logger.MoveLevel(-1);
            }

            BibleCommon.Services.Logger.MoveLevel(-1);
        }

        private void RelinkPageComments(string sectionGroupId, string sectionId, string pageId, string pageName)
        {
            XmlNamespaceManager xnm;
            XDocument pageDocument = OneNoteUtils.GetXDocument(OneNoteProxy.Instance.GetPageContent(_oneNoteApp, pageId), out xnm);

            вот здес - VerseLinkManager.FindVerseLinkPageAndCreateIfNeeded(_oneNoteApp, sectionId, pageId, pageName, SettingsManager.Instance.PageName_DefaultComments);
            надо искать с учётом имени страницы. (в ссылке искать строку, от ".one#" до "."). и при этом, чтобы одно искать то же не искать - запоминать где нить (особенно дефолтную страницу, потому что в 99,99% будет только она)

            foreach (XElement rowElement in pageDocument.Root.XPathSelectElements("one:Outline/one:OEChildren/one:OE/one:Table/one:Row/one:Cell[1]", xnm))
            {
                rowElement.Value = rowElement.Value.Replace("\n", " ");

                int linkIndex = rowElement.Value.IndexOf("<a ");

                while (linkIndex > -1)
                {
                    int linkEnd = rowElement.Value.IndexOf("</a>", linkIndex + 1);

                    if (linkEnd != -1)
                    {
                        RelinkPageComment(rowElement, linkIndex, linkEnd);                        

                        //ProgressBar.Progress()
                    }

                    linkIndex = rowElement.Value.IndexOf("<a ", linkIndex + 1);
                }
            }
        }

        private void RelinkPageComment(XElement rowElement, int linkIndex, int linkEnd)
        {
            
            string commentLink = rowElement.Value.Substring(linkIndex, linkEnd - linkIndex + "</a>".Length);
            string commentText = GetLinkText(commentLink);            
        }

        private string GetLinkText(string commentLink)
        {
            int breakIndex;
            string s = StringUtils.GetNextString(commentLink, -1, null, out breakIndex, StringSearchIgnorance.None,
                 StringSearchMode.SearchText);

            return s;
        }
       
    }
}
