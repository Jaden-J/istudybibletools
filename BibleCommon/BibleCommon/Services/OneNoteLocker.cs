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
        public static void LockBible(ref Application oneNoteApp)
        {
            string bibleHierarchyId = string.IsNullOrEmpty(SettingsManager.Instance.SectionGroupId_Bible)
                    ? SettingsManager.Instance.NotebookId_Bible : SettingsManager.Instance.SectionGroupId_Bible;

            LockOrUnlockHierarchy(ref oneNoteApp, bibleHierarchyId, true);
        }

        public static void LockSupplementalBible(ref Application oneNoteApp)
        {
            string supplementalBibleId = SettingsManager.Instance.GetValidSupplementalBibleNotebookId(ref oneNoteApp);

            if (!string.IsNullOrEmpty(supplementalBibleId))
                LockOrUnlockHierarchy(ref oneNoteApp, supplementalBibleId, true);
        }


        public static void UnlockBible(ref Application oneNoteApp)
        {
            string bibleHierarchyId = string.IsNullOrEmpty(SettingsManager.Instance.SectionGroupId_Bible)
                    ? SettingsManager.Instance.NotebookId_Bible : SettingsManager.Instance.SectionGroupId_Bible;

            LockOrUnlockHierarchy(ref oneNoteApp, bibleHierarchyId, false);
        }

        public static void UnlockSupplementalBible(ref Application oneNoteApp)
        {
            string supplementalBibleId = SettingsManager.Instance.GetValidSupplementalBibleNotebookId(ref oneNoteApp);

            if (!string.IsNullOrEmpty(supplementalBibleId))
                LockOrUnlockHierarchy(ref oneNoteApp, supplementalBibleId, false);
        }

        public static void LockCurrentSection(ref Application oneNoteApp)
        {
            var currentPageInfo = OneNoteUtils.GetCurrentPageInfo(ref oneNoteApp);

            OneNoteUtils.UseOneNoteAPI(ref oneNoteApp, (oneNoteAppSafe) =>
            {
                oneNoteAppSafe.SyncHierarchy(currentPageInfo.SectionId);
            });

            string sectionFilePath = GetElementPath(ref oneNoteApp, currentPageInfo.SectionId);

            LockSection(sectionFilePath);

            OneNoteUtils.UseOneNoteAPI(ref oneNoteApp, (oneNoteAppSafe) =>
            {
                oneNoteAppSafe.SyncHierarchy(currentPageInfo.SectionId);
            });
        }

        public static void UnlockCurrentSection(ref Application oneNoteApp)
        {
            var currentPageInfo = OneNoteUtils.GetCurrentPageInfo(ref oneNoteApp);

            string sectionFilePath = GetElementPath(ref oneNoteApp, currentPageInfo.SectionId);

            UnlockSection(sectionFilePath);

            OneNoteUtils.UseOneNoteAPI(ref oneNoteApp, (oneNoteAppSafe) =>
            {
                oneNoteAppSafe.SyncHierarchy(currentPageInfo.SectionId);
            });
        }

        private static void LockOrUnlockHierarchy(ref Application oneNoteApp, string hierarchyId, bool toLock)
        {
            if (SettingsManager.Instance.IsConfigured(ref oneNoteApp))
            {
                string folderPath = GetElementPath(ref oneNoteApp, hierarchyId);

                foreach (var filePath in Directory.GetFiles(folderPath, "*", SearchOption.AllDirectories))  // will throw NotSupportedException if Bible is stored on SkyDrive
                {
                    if (toLock)
                        LockSection(filePath);
                    else        
                        UnlockSection(filePath);
                }

                OneNoteUtils.UseOneNoteAPI(ref oneNoteApp, (oneNoteAppSafe) =>
                {
                    oneNoteAppSafe.SyncHierarchy(hierarchyId);
                });
            }
        }

        private static string GetElementPath(ref Application oneNoteApp, string elementId)
        {
            XmlNamespaceManager xnm;
            var xDoc = OneNoteUtils.GetHierarchyElement(ref oneNoteApp, elementId, HierarchyScope.hsSelf, out xnm);
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
