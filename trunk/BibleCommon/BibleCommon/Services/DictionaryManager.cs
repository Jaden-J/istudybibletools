using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.OneNote;
using BibleCommon.Helpers;
using System.Xml.Linq;
using System.IO;
using BibleCommon.Common;
using System.Xml;
using System.Threading;
using System.Xml.XPath;
using System.Runtime.InteropServices;

namespace BibleCommon.Services
{
    public static class DictionaryManager
    {
        public static void AddDictionary(Application oneNoteApp, string moduleName, string notebookDirectory, bool waitForFinish)
        {
            if (string.IsNullOrEmpty(SettingsManager.Instance.GetValidDictionariesNotebookId(oneNoteApp, true)))
            {
                SettingsManager.Instance.NotebookId_Dictionaries
                    = NotebookGenerator.CreateNotebook(oneNoteApp, Resources.Constants.DictionariesNotebookName, notebookDirectory);
                SettingsManager.Instance.DictionariesModules.Clear();
                SettingsManager.Instance.Save();
            }

            if (!SettingsManager.Instance.DictionariesModules.Any(m => m.ModuleName == moduleName))
            {
                //section or sectionGroup Id
                string dictionarySectionId = null;
                string dictionarySectionPath = null;
                XElement dictionarySectionEl = null;

                var moduleInfo = ModulesManager.GetModuleInfo(moduleName);

                if (moduleInfo.NotebooksStructure.Sections == null || moduleInfo.NotebooksStructure.Sections.Count == 0)
                    throw new InvalidModuleException("There is no information about dictionary sections.");

                if (string.IsNullOrEmpty(moduleInfo.NotebooksStructure.DictionarySectionGroupName) && moduleInfo.NotebooksStructure.Sections.Count > 1)
                    moduleInfo.NotebooksStructure.DictionarySectionGroupName = moduleInfo.Name;

                if (!string.IsNullOrEmpty(moduleInfo.NotebooksStructure.DictionarySectionGroupName))
                {
                    dictionarySectionEl = NotebookGenerator.AddRootSectionGroupToNotebook(oneNoteApp, SettingsManager.Instance.NotebookId_Dictionaries, moduleInfo.NotebooksStructure.DictionarySectionGroupName);                                        
                }
                else
                {
                    XmlNamespaceManager xnm;
                    dictionarySectionEl = OneNoteUtils.GetHierarchyElement(oneNoteApp, SettingsManager.Instance.NotebookId_Dictionaries, HierarchyScope.hsSelf, out xnm).Root;
                }

                dictionarySectionId = (string)dictionarySectionEl.Attribute("ID");
                dictionarySectionPath = (string)dictionarySectionEl.Attribute("path");
                
                oneNoteApp.SyncHierarchy(dictionarySectionId);
                while (!Directory.Exists(dictionarySectionPath))
                {
                    Thread.Sleep(1000);
                    System.Windows.Forms.Application.DoEvents();
                }

                foreach (var sectionInfo in moduleInfo.NotebooksStructure.Sections)
                {
                    string sectionElId;                    

                    File.Copy(Path.Combine(ModulesManager.GetModuleDirectory(moduleName), sectionInfo.Name), Path.Combine(dictionarySectionPath, sectionInfo.Name), false);
                    oneNoteApp.OpenHierarchy(sectionInfo.Name, dictionarySectionId, out sectionElId, CreateFileType.cftSection);

                    if (string.IsNullOrEmpty(dictionarySectionId))
                        dictionarySectionId = sectionElId;
                }

                oneNoteApp.SyncHierarchy(dictionarySectionId);

                SettingsManager.Instance.DictionariesModules.Add(new DictionaryInfo(moduleName, dictionarySectionId));
                SettingsManager.Instance.Save();

                if (waitForFinish)                
                    WaitWhileDictionaryIsCreating(oneNoteApp, dictionarySectionId, moduleInfo.NotebooksStructure.DictionaryPagesCount, 0);                
            }
        }

        private static void WaitWhileDictionaryIsCreating(Application oneNoteApp, string dictionarySectionId, int? dictionaryPagesCount, int attemptsCount)
        {
            if (dictionaryPagesCount.HasValue)
            {
                if (attemptsCount < 1000)
                {
                    XmlNamespaceManager xnm;
                    var xDoc = OneNoteUtils.GetHierarchyElement(oneNoteApp, dictionarySectionId, HierarchyScope.hsPages, out xnm);
                    int pagesCount = xDoc.Root.XPathSelectElements("//one:Page", xnm).Count();
                    if (pagesCount < dictionaryPagesCount)
                    {
                        Thread.Sleep(3000);
                        WaitWhileDictionaryIsCreating(oneNoteApp, dictionarySectionId, dictionaryPagesCount, attemptsCount + 1);
                    }
                }
            }
        }

        public static void RemoveDictionary(Application oneNoteApp, string moduleName)
        {
            if (SettingsManager.Instance.DictionariesModules.Count == 1 
                && SettingsManager.Instance.DictionariesModules.First().ModuleName == moduleName)
            {
                CloseDictionariesNotebook(oneNoteApp);
            }
            else
            {
                var dictionaryModuleInfo = SettingsManager.Instance.DictionariesModules.FirstOrDefault(m => m.ModuleName == moduleName);
                if (dictionaryModuleInfo != null)
                {
                    try
                    {
                        oneNoteApp.DeleteHierarchy(dictionaryModuleInfo.SectionId);
                    }
                    catch (COMException ex)
                    {
                        if (!ex.Message.Contains("0x80042014"))   // The object does not exist.
                            throw;
                    }
                    SettingsManager.Instance.DictionariesModules.Remove(dictionaryModuleInfo);
                    SettingsManager.Instance.Save();
                }
            }
        }

        public static void CloseDictionariesNotebook(Application oneNoteApp)
        {
            OneNoteUtils.CloseNotebookSafe(oneNoteApp, SettingsManager.Instance.NotebookId_Dictionaries);

            SettingsManager.Instance.NotebookId_Dictionaries = null;
            SettingsManager.Instance.DictionariesModules.Clear();
            SettingsManager.Instance.Save();
        }
    }
}
