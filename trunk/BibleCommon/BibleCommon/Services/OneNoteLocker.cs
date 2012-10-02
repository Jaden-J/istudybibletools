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
using Microsoft.Office.Interop.OneNote;

namespace BibleCommon.Services
{
    public static class OneNoteLocker
    {
        public static void LockBible(Application oneNoteApp)
        {
            string bibleHierarchyId = string.IsNullOrEmpty(SettingsManager.Instance.SectionGroupId_Bible)
                    ? SettingsManager.Instance.NotebookId_Bible : SettingsManager.Instance.SectionGroupId_Bible;

            LockOrUnlockHierarchy(oneNoteApp, bibleHierarchyId, true);
        }

        public static void LockSupplementalBible(Application oneNoteApp)
        {
            string supplementalBibleId = SettingsManager.Instance.GetValidSupplementalBibleNotebookId(oneNoteApp);

            if (!string.IsNullOrEmpty(supplementalBibleId))
                LockOrUnlockHierarchy(oneNoteApp, supplementalBibleId, true);
        }


        public static void UnlockBible(Application oneNoteApp)
        {
            string bibleHierarchyId = string.IsNullOrEmpty(SettingsManager.Instance.SectionGroupId_Bible)
                    ? SettingsManager.Instance.NotebookId_Bible : SettingsManager.Instance.SectionGroupId_Bible;

            LockOrUnlockHierarchy(oneNoteApp, bibleHierarchyId, false);
        }

        public static void UnlockSupplementalBible(Application oneNoteApp)
        {
            string supplementalBibleId = SettingsManager.Instance.GetValidSupplementalBibleNotebookId(oneNoteApp);

            if (!string.IsNullOrEmpty(supplementalBibleId))
                LockOrUnlockHierarchy(oneNoteApp, supplementalBibleId, false);
        }

        public static void LockCurrentSection(Application oneNoteApp)
        {
            var currentPageInfo = OneNoteUtils.GetCurrentPageInfo(oneNoteApp);

            oneNoteApp.SyncHierarchy(currentPageInfo.SectionId);

            string sectionFilePath = GetElementPath(oneNoteApp, currentPageInfo.SectionId);

            LockSection(sectionFilePath);

            oneNoteApp.SyncHierarchy(currentPageInfo.SectionId);
        }

        public static void UnlockCurrentSection(Application oneNoteApp)
        {
            var currentPageInfo = OneNoteUtils.GetCurrentPageInfo(oneNoteApp);

            string sectionFilePath = GetElementPath(oneNoteApp, currentPageInfo.SectionId);

            UnlockSection(sectionFilePath);

            oneNoteApp.SyncHierarchy(currentPageInfo.SectionId);
        }

        private static void LockOrUnlockHierarchy(Application oneNoteApp, string hierarchyId, bool toLock)
        {
            if (SettingsManager.Instance.IsConfigured(oneNoteApp))
            {
                string folderPath = GetElementPath(oneNoteApp, hierarchyId);

                foreach (var filePath in Directory.GetFiles(folderPath, "*", SearchOption.AllDirectories))  // will throw NotSupportedException if Bible is stored on SkyDrive
                {
                    if (toLock)
                        LockSection(filePath);
                    else        
                        UnlockSection(filePath);
                }

                oneNoteApp.SyncHierarchy(hierarchyId);
            }
        }

        private static string GetElementPath(Application oneNoteApp, string elementId)
        {
            XmlNamespaceManager xnm;
            var xDoc = OneNoteUtils.GetHierarchyElement(oneNoteApp, elementId, HierarchyScope.hsSelf, out xnm);
            return (string)xDoc.Root.Attribute("path");
        }

        private static void LockSection(string sectionFilePath)
        {
            File.SetAttributes(sectionFilePath, FileAttributes.ReadOnly);  // will throw NotSupportedException if Bible is stored in SkyDrive
        } 

        private static void UnlockSection(string sectionFilePath)
        {
            File.SetAttributes(sectionFilePath, FileAttributes.Normal);    // will throw NotSupportedException if Bible is stored in SkyDrive         
        }
    }
}
