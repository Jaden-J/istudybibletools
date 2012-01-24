using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BibleCommon;
using Microsoft.Office.Interop.OneNote;
using System.Configuration;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;
using BibleCommon.Services;
using BibleCommon.Helpers;
using BibleCommon.Consts;

namespace BibleTablesResizer
{
    class Program
    {
        static void Main(string[] args)
        {
            Logger.Init("BibleTablesResizer");

            Logger.LogMessage("Старт обработки всей записной книжки");

            Application oneNoteApp = new Application();

            string notebookId = SettingsManager.Instance.NotebookId_Bible;            

            if (!string.IsNullOrEmpty(notebookId))
            {
                Logger.LogMessage("Имя записной книжки: {0}", OneNoteUtils.GetHierarchyElementName(oneNoteApp, notebookId));  // чтобы точно убедиться

                string hierarchyXml;
                oneNoteApp.GetHierarchy(notebookId, HierarchyScope.hsPages, out hierarchyXml);
                XmlNamespaceManager xnm;
                XDocument notebookDoc = OneNoteUtils.GetXDocument(hierarchyXml, out xnm);

                ProcessRootSectionGroup(oneNoteApp, notebookId, notebookDoc, SettingsManager.Instance.SectionGroupId_Bible, xnm);
            }
            else
            {
                Logger.LogError(string.Format("Не найдено записной книжки '{0}'", notebookId));
            }

            Logger.Done();

            if (Logger.ErrorWasLogged)
            {
                Console.WriteLine("Во время работы программы произошли ошибки.");
                Console.ReadKey();
            }
        }

        private static string GetBibleNotebookName(Application oneNoteApp, string notebookName)
        {
            string xml;
            XmlNamespaceManager xnm;
            oneNoteApp.GetHierarchy(null, HierarchyScope.hsNotebooks, out xml);
            XDocument doc = OneNoteUtils.GetXDocument(xml, out xnm);
            XElement bibleNotebook = doc.Root.XPathSelectElement(string.Format("one:Notebook[@ID='{0}']", notebookName), xnm);
            if (bibleNotebook != null)
            {
                return (string)bibleNotebook.Attribute("name");
            }

            return null;
        }

        private static void ProcessRootSectionGroup(Application oneNoteApp, string notebookId, XDocument doc, string sectionGroupId, XmlNamespaceManager xnm)
        {
            XElement sectionGroup = doc.Root.XPathSelectElement(
                            string.Format("one:SectionGroup[@ID='{0}']", sectionGroupId), xnm);

            if (sectionGroup != null)
                ProcessSectionGroup(sectionGroup, oneNoteApp, notebookId, xnm);
            else
                Logger.LogError("Не удаётся найти группу секций '{0}'", sectionGroupId);
        }

        private static void ProcessSectionGroup(XElement sectionGroup,
            Application oneNoteApp, string notebookId, XmlNamespaceManager xnm)
        {
            string sectionGroupId = (string)sectionGroup.Attribute("ID");
            string sectionGroupName = (string)sectionGroup.Attribute("name");

            Logger.LogMessage("Обработка группы секций '{0}'", sectionGroupName);
            Logger.MoveLevel(1);

            foreach (var subSectionGroup in sectionGroup.XPathSelectElements("one:SectionGroup", xnm))
            {
                ProcessSectionGroup(subSectionGroup, oneNoteApp, notebookId, xnm);
            }

            foreach (var subSection in sectionGroup.XPathSelectElements("one:Section", xnm))
            {
                ProcessSection(subSection, sectionGroupId, oneNoteApp, notebookId, xnm);
            }

            Logger.MoveLevel(-1);
        }

        private static void ProcessSection(XElement section, string sectionGroupId,
            Application oneNoteApp, string notebookId, XmlNamespaceManager xnm)
        {
            string sectionId = (string)section.Attribute("ID");
            string sectionName = (string)section.Attribute("name");

            Logger.LogMessage("Обработка секции '{0}'", sectionName);
            Logger.MoveLevel(1);

            foreach (var page in section.XPathSelectElements("one:Page", xnm))
            {
                string pageId = (string)page.Attribute("ID");
                string pageName = (string)page.Attribute("name");

                Logger.LogMessage("Обработка страницы '{0}'", pageName);

                Logger.MoveLevel(1);

                TableModifier.ModifyTable(oneNoteApp, pageId);                

                Logger.MoveLevel(-1);
            }

            Logger.MoveLevel(-1);
        }        
    }
}
