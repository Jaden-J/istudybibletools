using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BibleCommon.Helpers;
using Microsoft.Office.Interop.OneNote;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;

namespace BibleConfigurator.ModuleConverter
{
    public static class NotebookGenerator
    {
        public static void GenerateSummaryOfNotesNotebook(string bibleNotebookName, string targetEmptyNotebookName)
        {
            var oneNoteApp = new Application();

            string bibleNotebookId = OneNoteUtils.GetNotebookIdByName(oneNoteApp, bibleNotebookName, false);
            XmlNamespaceManager xnm;
            var bibleNotebookDoc = OneNoteUtils.GetHierarchyElement(oneNoteApp, bibleNotebookId, HierarchyScope.hsSections, out xnm);

            string targetNotebookId = OneNoteUtils.GetNotebookIdByName(oneNoteApp, targetEmptyNotebookName, false);                        

            foreach (var testamentSectionGroup in bibleNotebookDoc.Root.XPathSelectElements("one:SectionGroup", xnm))
            {
                string testamentSectionGroupName = testamentSectionGroup.Attribute("name").Value;
                XElement testamentSectionGroupEl =  OneNoteUtils.AddRootSectionGroupToNotebook(oneNoteApp, targetNotebookId, testamentSectionGroupName);                

                foreach (var bibleBookSection in testamentSectionGroup.XPathSelectElements("one:Section", xnm))
                {
                    string bibleBookSectionName = bibleBookSection.Attribute("name").Value;
                    OneNoteUtils.AddSectionGroup(oneNoteApp, testamentSectionGroupEl, bibleBookSectionName);
                }
            }
        }
    }
}
