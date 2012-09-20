using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.OneNote;
using BibleCommon.Helpers;
using System.Xml.Linq;
using System.IO;
using BibleCommon.Common;

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

                var moduleInfo = ModulesManager.GetModuleInfo(moduleName);

                if (moduleInfo.DictionarySections == null || moduleInfo.DictionarySections.Count == 0)
                    throw new InvalidModuleException("There is no information about dictionary sections.");

                if (string.IsNullOrEmpty(moduleInfo.DictionarySectionGroupName) && moduleInfo.DictionarySections.Count > 1)
                    moduleInfo.DictionarySectionGroupName = moduleInfo.Name;

                if (!string.IsNullOrEmpty(moduleInfo.DictionarySectionGroupName))
                {
                    var sectionGroupEl = NotebookGenerator.AddRootSectionGroupToNotebook(oneNoteApp, SettingsManager.Instance.NotebookId_Dictionaries, moduleInfo.DictionarySectionGroupName);
                    dictionarySectionId = (string)sectionGroupEl.Attribute("ID");
                }

                foreach(var sectionInfo in moduleInfo.DictionarySections)
                {
                    string sectionElId;
                    oneNoteApp.OpenHierarchy(Path.Combine(ModulesManager.GetModuleDirectory(moduleName), sectionInfo.Name), 
                                                dictionarySectionId ?? SettingsManager.Instance.NotebookId_Dictionaries, out sectionElId);

                    if (string.IsNullOrEmpty(dictionarySectionId))
                        dictionarySectionId = sectionElId;
                }

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
            oneNoteApp.CloseNotebook(SettingsManager.Instance.NotebookId_Dictionaries);

            SettingsManager.Instance.NotebookId_Dictionaries = null;
            SettingsManager.Instance.DictionariesModules.Clear();
            SettingsManager.Instance.Save();
        }
    }
}
