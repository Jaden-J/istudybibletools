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
using BibleCommon.Contracts;
using BibleCommon.Common;

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


        public static void UnlockBible(ref Application oneNoteApp, bool wait = false, Func<bool> checkIfExternalProcessAborted = null)
        {
            string bibleHierarchyId = string.IsNullOrEmpty(SettingsManager.Instance.SectionGroupId_Bible)
                    ? SettingsManager.Instance.NotebookId_Bible : SettingsManager.Instance.SectionGroupId_Bible;

            LockOrUnlockHierarchy(ref oneNoteApp, bibleHierarchyId, false);

            if (wait)
                WaitForHierarchyIsUnlocked(ref oneNoteApp, bibleHierarchyId, 0, checkIfExternalProcessAborted);            
        }

        public static void UnlockSupplementalBible(ref Application oneNoteApp, bool wait = false, Func<bool> checkIfExternalProcessAborted = null)
        {
            string supplementalBibleId = SettingsManager.Instance.GetValidSupplementalBibleNotebookId(ref oneNoteApp);

            if (!string.IsNullOrEmpty(supplementalBibleId))
                LockOrUnlockHierarchy(ref oneNoteApp, supplementalBibleId, false);

            if (wait)
                WaitForHierarchyIsUnlocked(ref oneNoteApp, supplementalBibleId, 0, checkIfExternalProcessAborted);
        }

        public static void LockCurrentSection(ref Application oneNoteApp)
        {
            var currentPageInfo = OneNoteUtils.GetCurrentPageInfo(ref oneNoteApp);

            OneNoteUtils.UseOneNoteAPI(ref oneNoteApp, (oneNoteAppSafe) =>
            {
                oneNoteAppSafe.SyncHierarchy(currentPageInfo.SectionId);
            });

            string sectionFilePath = OneNoteUtils.GetElementPath(ref oneNoteApp, currentPageInfo.SectionId);

            LockSection(sectionFilePath);

            OneNoteUtils.UseOneNoteAPI(ref oneNoteApp, (oneNoteAppSafe) =>
            {
                oneNoteAppSafe.SyncHierarchy(currentPageInfo.SectionId);
            });
        }

        public static void UnlockCurrentSection(ref Application oneNoteApp)
        {
            var currentPageInfo = OneNoteUtils.GetCurrentPageInfo(ref oneNoteApp);

            string sectionFilePath = OneNoteUtils.GetElementPath(ref oneNoteApp, currentPageInfo.SectionId);

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
                string folderPath = OneNoteUtils.GetElementPath(ref oneNoteApp, hierarchyId);

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

       

        private static void LockSection(string sectionFilePath)
        {
            File.SetAttributes(sectionFilePath, FileAttributes.ReadOnly);  // will throw NotSupportedException if Bible is stored in SkyDrive
        } 

        private static void UnlockSection(string sectionFilePath)
        {
            File.SetAttributes(sectionFilePath, FileAttributes.Normal);    // will throw NotSupportedException if Bible is stored in SkyDrive         
        }

        public static void WaitForHierarchyIsUnlocked(ref Application oneNoteApp, string hierarchyId, int attemptsCount, Func<bool> checkIfExternalProcessAborted)
        {
            string folderPath = OneNoteUtils.GetElementPath(ref oneNoteApp, hierarchyId);

            if (attemptsCount < 10)   // 30 секунд
            {

                try
                {
                    var success = true;
                    foreach (var filePath in Directory.GetFiles(folderPath, "*", SearchOption.AllDirectories))
                    {
                        if ((File.GetAttributes(filePath) & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
                        {
                            success = false;
                            break;
                        }
                    }

                    if (!success)
                    {
                        Utils.WaitFor3Seconds(checkIfExternalProcessAborted);
                        WaitForHierarchyIsUnlocked(ref oneNoteApp, hierarchyId, attemptsCount + 1, checkIfExternalProcessAborted);
                    }
                }
                catch (NotSupportedException)
                {
                }                
            }
        }
    }
}
