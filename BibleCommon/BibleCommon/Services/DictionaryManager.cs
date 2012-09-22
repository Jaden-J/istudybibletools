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

namespace BibleCommon.Services
{
    public static class DictionaryManager
    {
        public static void AddDictionary(Application oneNoteApp, string moduleName, string notebookDirectory)
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

                if (moduleInfo.DictionarySections == null || moduleInfo.DictionarySections.Count == 0)
                    throw new InvalidModuleException("There is no information about dictionary sections.");

                if (string.IsNullOrEmpty(moduleInfo.DictionarySectionGroupName) && moduleInfo.DictionarySections.Count > 1)
                    moduleInfo.DictionarySectionGroupName = moduleInfo.Name;

                if (!string.IsNullOrEmpty(moduleInfo.DictionarySectionGroupName))
                {
                    dictionarySectionEl = NotebookGenerator.AddRootSectionGroupToNotebook(oneNoteApp, SettingsManager.Instance.NotebookId_Dictionaries, moduleInfo.DictionarySectionGroupName);                                        
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
                }

                foreach(var sectionInfo in moduleInfo.DictionarySections)
                {
                    string sectionElId;                    

                    File.Copy(Path.Combine(ModulesManager.GetModuleDirectory(moduleName), sectionInfo.Name), Path.Combine(dictionarySectionPath, sectionInfo.Name), false);
                    oneNoteApp.OpenHierarchy(sectionInfo.Name, dictionarySectionId, out sectionElId, CreateFileType.cftSection);

                    if (string.IsNullOrEmpty(dictionarySectionId))
                        dictionarySectionId = sectionElId;
                }

                oneNoteApp.SyncHierarchy(dictionarySectionId);

                SettingsManager.Instance.DictionariesModules.Add(new DictionaryModuleInfo(moduleName, dictionarySectionId));
                SettingsManager.Instance.Save();
            }
        }

        public static void RemoveDictionary(Application oneNoteApp, string moduleName)
        {
            if (SettingsManager.Instance.DictionariesModules.Count == 1)
            {
                CloseDictionariesNotebook(oneNoteApp);
            }
            else
            {
                var dictionaryModuleInfo = SettingsManager.Instance.DictionariesModules.FirstOrDefault(m => m.ModuleName == moduleName);
                if (dictionaryModuleInfo != null)
                {
                    oneNoteApp.DeleteHierarchy(dictionaryModuleInfo.SectionId);
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
