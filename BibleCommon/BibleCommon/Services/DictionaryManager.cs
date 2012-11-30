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
using BibleCommon.UI.Forms;

namespace BibleCommon.Services
{
    public static class DictionaryManager
    {
        public static void AddDictionary(Application oneNoteApp, ModuleInfo module, string notebookDirectory, bool waitForFinish)
        {
            if (string.IsNullOrEmpty(SettingsManager.Instance.GetValidDictionariesNotebookId(oneNoteApp, true)))
            {
                SettingsManager.Instance.NotebookId_Dictionaries
                    = NotebookGenerator.CreateNotebook(oneNoteApp, Resources.Constants.DictionariesNotebookName, notebookDirectory);
                SettingsManager.Instance.DictionariesModules.Clear();
                SettingsManager.Instance.Save();
            }

            if (!SettingsManager.Instance.DictionariesModules.Any(m => m.ModuleName == module.ShortName))
            {
                //section or sectionGroup Id
                string dictionarySectionId = null;
                string dictionarySectionPath = null;
                XElement dictionarySectionEl = null;

                var moduleInfo = ModulesManager.GetModuleInfo(module.ShortName);

                if (moduleInfo.NotebooksStructure.Sections == null || moduleInfo.NotebooksStructure.Sections.Count == 0)
                    throw new InvalidModuleException("There is no information about dictionary sections.");

                if (string.IsNullOrEmpty(moduleInfo.NotebooksStructure.DictionarySectionGroupName) && moduleInfo.NotebooksStructure.Sections.Count > 1)
                    moduleInfo.NotebooksStructure.DictionarySectionGroupName = moduleInfo.DisplayName;

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
                int attemptsCount = 0;
                while (!Directory.Exists(dictionarySectionPath) && attemptsCount < 500)
                {
                    Thread.Sleep(1000);
                    System.Windows.Forms.Application.DoEvents();
                    attemptsCount++;
                }

                foreach (var sectionInfo in moduleInfo.NotebooksStructure.Sections)
                {
                    string sectionElId;                    

                    File.Copy(Path.Combine(ModulesManager.GetModuleDirectory(module.ShortName), sectionInfo.Name), Path.Combine(dictionarySectionPath, sectionInfo.Name), false);
                    oneNoteApp.OpenHierarchy(sectionInfo.Name, dictionarySectionId, out sectionElId, CreateFileType.cftSection);

                    if (string.IsNullOrEmpty(dictionarySectionId))
                        dictionarySectionId = sectionElId;
                }

                oneNoteApp.SyncHierarchy(dictionarySectionId);

                SettingsManager.Instance.DictionariesModules.Add(new StoredModuleInfo(module.ShortName, module.Version, dictionarySectionId));
                SettingsManager.Instance.Save();

                if (waitForFinish)                
                    WaitWhileDictionaryIsCreating(oneNoteApp, dictionarySectionId, moduleInfo.NotebooksStructure.DictionaryPagesCount, 0);                
            }
        }

        public static void WaitWhileDictionaryIsCreating(Application oneNoteApp, string dictionarySectionId, int? dictionaryPagesCount, int attemptsCount)
        {
            if (dictionaryPagesCount.HasValue)
            {
                if (attemptsCount < 500)  // 25 минут
                {
                    XmlNamespaceManager xnm;
                    var xDoc = OneNoteUtils.GetHierarchyElement(oneNoteApp, dictionarySectionId, HierarchyScope.hsPages, out xnm);
                    int pagesCount = xDoc.Root.XPathSelectElements("//one:Page", xnm).Count();
                    if (pagesCount < dictionaryPagesCount)
                    {
                        oneNoteApp.SyncHierarchy(dictionarySectionId);
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
                        if (!ex.Message.Contains(Utils.GetHexError(Error.hrObjectDoesNotExist)))
                            throw;
                    }

                    DictionaryTermsCacheManager.RemoveCache(dictionaryModuleInfo.ModuleName);

                    SettingsManager.Instance.DictionariesModules.Remove(dictionaryModuleInfo);
                    SettingsManager.Instance.Save();
                }
            }
        }

        public static void CloseDictionariesNotebook(Application oneNoteApp)
        {
            OneNoteUtils.CloseNotebookSafe(oneNoteApp, SettingsManager.Instance.NotebookId_Dictionaries);

            SettingsManager.Instance.NotebookId_Dictionaries = null;

            foreach (var dictionaryInfo in SettingsManager.Instance.DictionariesModules)
            {
                DictionaryTermsCacheManager.RemoveCache(dictionaryInfo.ModuleName);
            }

            SettingsManager.Instance.DictionariesModules.Clear();
            SettingsManager.Instance.Save();
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="oneNoteApp"></param>
        /// <param name="link"></param>
        /// <returns>если false - надо перестраивать кэш</returns>
        public static bool GoToTerm(Application oneNoteApp, DictionaryTermLink link)
        {
            try
            {
                oneNoteApp.NavigateTo(link.PageId, link.ObjectId);
                return true;
            }
            catch (COMException ex)
            {
                if (ex.Message.Contains(Utils.GetHexError(Error.hrObjectDoesNotExist)))
                {
                    using (var form = new MessageForm(BibleCommon.Resources.Constants.RebuldDictionaryCache, BibleCommon.Resources.Constants.Warning,
                            System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question))
                    {
                        if (form.ShowDialog() == System.Windows.Forms.DialogResult.Yes)                        
                            return false;                        
                        else
                            return true;
                    }
                }
                else
                    throw;
            }
        }
    }
}
