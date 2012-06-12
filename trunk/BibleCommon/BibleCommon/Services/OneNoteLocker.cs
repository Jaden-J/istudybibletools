using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using BibleCommon.Helpers;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;
using System.Threading;

namespace BibleCommon.Services
{
    public static class OneNoteLocker
    {
        public static void LockAllBible(Microsoft.Office.Interop.OneNote.Application oneNoteApp)
        {
            LockOrUnlockAllBIble(oneNoteApp, true);
        }

        public static void UnlockAllBible(Microsoft.Office.Interop.OneNote.Application oneNoteApp)
        {
            LockOrUnlockAllBIble(oneNoteApp, false);
        }

        public static void UnlockCurrentSection(Microsoft.Office.Interop.OneNote.Application oneNoteApp)
        {
            var currentPageInfo = OneNoteUtils.GetCurrentPageInfo(oneNoteApp);

            string sectionFilePath = GetElementPath(oneNoteApp, currentPageInfo.SectionId);

            UnlockSection(sectionFilePath);

            oneNoteApp.SyncHierarchy(currentPageInfo.SectionId);
        }

        private static void LockOrUnlockAllBIble(Microsoft.Office.Interop.OneNote.Application oneNoteApp, bool toLock)
        {
            if (SettingsManager.Instance.IsConfigured(oneNoteApp))
            {
                string hierarchyId = string.IsNullOrEmpty(SettingsManager.Instance.SectionGroupId_Bible)
                    ? SettingsManager.Instance.NotebookId_Bible : SettingsManager.Instance.SectionGroupId_Bible;

                string folderPath = GetElementPath(oneNoteApp, hierarchyId);

                foreach (var filePath in Directory.GetFiles(folderPath, "*", SearchOption.AllDirectories))
                {
                    if (toLock)
                        LockSection(filePath);
                    else        
                        UnlockSection(filePath);
                }

                oneNoteApp.SyncHierarchy(hierarchyId);
            }
        }

        private static string GetElementPath(Microsoft.Office.Interop.OneNote.Application oneNoteApp, string elementId)
        {
            XmlNamespaceManager xnm;
            var xDoc = OneNoteUtils.GetHierarchyElement(oneNoteApp, elementId, Microsoft.Office.Interop.OneNote.HierarchyScope.hsSelf, out xnm);
            return xDoc.Root.Attribute("path").Value;
        }

        private static void LockSection(string sectionFilePath)
        {
            File.SetAttributes(sectionFilePath, FileAttributes.ReadOnly);
        }

        private static void UnlockSection(string sectionFilePath)
        {
            File.SetAttributes(sectionFilePath, FileAttributes.Normal);            
        }
    }
}
