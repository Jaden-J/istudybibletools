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
        public static void AddDictionary(ref Application oneNoteApp, ModuleInfo moduleInfo, string notebookDirectory, bool waitForFinish, Func<bool> checkIfExternalProcessAborted)
        {
            if (string.IsNullOrEmpty(SettingsManager.Instance.GetValidDictionariesNotebookId(ref oneNoteApp, true)))
            {
                SettingsManager.Instance.NotebookId_Dictionaries
                    = NotebookGenerator.CreateNotebook(ref oneNoteApp, Resources.Constants.DictionariesNotebookName, notebookDirectory, Resources.Constants.DictionariesNotebookName);
                SettingsManager.Instance.DictionariesModules.Clear();
                SettingsManager.Instance.Save();
            }

            var existingDictionaryModuleInfo = SettingsManager.Instance.DictionariesModules.FirstOrDefault(m => m.ModuleName == moduleInfo.ShortName);

            if (existingDictionaryModuleInfo == null 
                || !OneNoteUtils.HierarchyElementExists(ref oneNoteApp, existingDictionaryModuleInfo.SectionId))
            {
                if (existingDictionaryModuleInfo != null)
                    SettingsManager.Instance.DictionariesModules.Remove(existingDictionaryModuleInfo);

                //section or sectionGroup Id
                string dictionarySectionId = null;                

                if (moduleInfo.NotebooksStructure.Sections == null || moduleInfo.NotebooksStructure.Sections.Count == 0)
                    throw new InvalidModuleException("There is no information about dictionary sections.");

                if (string.IsNullOrEmpty(moduleInfo.NotebooksStructure.DictionarySectionGroupName) && moduleInfo.NotebooksStructure.Sections.Count > 1)
                    moduleInfo.NotebooksStructure.DictionarySectionGroupName = moduleInfo.DisplayName;

                dictionarySectionId = TryGetExistingDictionarySection(ref oneNoteApp, moduleInfo);

                if (string.IsNullOrEmpty(dictionarySectionId))
                {
                    dictionarySectionId = AddDictionaryToOneNote(ref oneNoteApp, moduleInfo, checkIfExternalProcessAborted);
                }           

                SettingsManager.Instance.DictionariesModules.Add(new StoredModuleInfo(moduleInfo.ShortName, moduleInfo.Version, dictionarySectionId));
                SettingsManager.Instance.Save();

                if (waitForFinish)                
                    WaitWhileDictionaryIsCreating(ref oneNoteApp, dictionarySectionId, moduleInfo.NotebooksStructure.DictionaryPagesCount, 0, checkIfExternalProcessAborted);                
            }
        }

        private static string TryGetExistingDictionarySection(ref Application oneNoteApp, ModuleInfo moduleInfo)
        {
            string dictionarySectionId = null;
            XElement dictionarySectionEl;
            XmlNamespaceManager xnm;

            if (!string.IsNullOrEmpty(moduleInfo.NotebooksStructure.DictionarySectionGroupName))
            {
                dictionarySectionEl = OneNoteUtils.GetHierarchyElementByName(ref oneNoteApp,
                                                        "SectionGroup", moduleInfo.NotebooksStructure.DictionarySectionGroupName,
                                                        SettingsManager.Instance.NotebookId_Dictionaries);
                if (dictionarySectionEl != null)
                {
                    dictionarySectionId = (string)dictionarySectionEl.Attribute("ID");
                    var sectionsDoc = OneNoteUtils.GetHierarchyElement(ref oneNoteApp, dictionarySectionId, HierarchyScope.hsSections, out xnm);
                    dictionarySectionEl = sectionsDoc.Root;
                }
            }
            else
            {
                dictionarySectionEl = OneNoteUtils.GetHierarchyElementByName(ref oneNoteApp,
                                                        "Section", moduleInfo.NotebooksStructure.Sections.First().Name,
                                                        SettingsManager.Instance.NotebookId_Dictionaries);
                if (dictionarySectionEl != null)                
                    dictionarySectionId = (string)dictionarySectionEl.Attribute("ID");
            }

            if (dictionarySectionEl != null) // dictionarySectionEl может быть как группа разделов, так и раздел. 
            {
                var pagesDoc = OneNoteUtils.GetHierarchyElement(ref oneNoteApp, (string)dictionarySectionEl.Attribute("ID"), HierarchyScope.hsPages, out xnm);                

                foreach (var pageEl in pagesDoc.Root.XPathSelectElements("//one:Page", xnm))
                {
                    var pageId = (string)pageEl.Attribute("ID");
                    var pageName = (string)pageEl.Attribute("name");

                    var embeddedModulesInfo_string = OneNoteUtils.GetPageMetaData(pageEl, BibleCommon.Consts.Constants.Key_EmbeddedDictionaries, xnm);
                    if (!string.IsNullOrEmpty(embeddedModulesInfo_string))
                    {
                        var embeddedModulesInfo = EmbeddedModuleInfo.Deserialize(embeddedModulesInfo_string);

                        foreach (var embeddedModuleInfo in embeddedModulesInfo)
                        {
                            if (moduleInfo.ShortName == embeddedModuleInfo.ModuleName)
                                return dictionarySectionId;
                        }
                    }
                }
            }

            return null; // иначе если группа секций/раздел и существует, то значит принадлежат другому модулю
        }

        private static string AddDictionaryToOneNote(ref Application oneNoteApp, ModuleInfo moduleInfo, Func<bool> checkIfExternalProcessAborted)
        {
            string dictionarySectionId;
            string dictionarySectionPath = null;
            XElement dictionarySectionEl = null;                
            
            if (!string.IsNullOrEmpty(moduleInfo.NotebooksStructure.DictionarySectionGroupName))
            {
                dictionarySectionEl = NotebookGenerator.AddRootSectionGroupToNotebook(ref oneNoteApp, SettingsManager.Instance.NotebookId_Dictionaries,
                                                            moduleInfo.NotebooksStructure.DictionarySectionGroupName, "." + moduleInfo.ShortName);
            }
            else
            {
                XmlNamespaceManager xnm;
                dictionarySectionEl = OneNoteUtils.GetHierarchyElement(ref oneNoteApp, SettingsManager.Instance.NotebookId_Dictionaries, HierarchyScope.hsSelf, out xnm).Root;
            }

            dictionarySectionId = (string)dictionarySectionEl.Attribute("ID");
            dictionarySectionPath = (string)dictionarySectionEl.Attribute("path");

            OneNoteUtils.UseOneNoteAPI(ref oneNoteApp, (oneNoteAppSafe) =>
            {
                oneNoteAppSafe.SyncHierarchy(dictionarySectionId);
            });

            int attemptsCount = 0;
            while (!Directory.Exists(dictionarySectionPath) && attemptsCount < 500)
            {
                Thread.Sleep(1000);

                if (checkIfExternalProcessAborted != null)
                {
                    if (checkIfExternalProcessAborted())
                        throw new ProcessAbortedByUserException();
                }

                System.Windows.Forms.Application.DoEvents();
                attemptsCount++;
            }

            foreach (var sectionInfo in moduleInfo.NotebooksStructure.Sections) //todo: если у нас нет отдельной группы разделов и такой уже раздел существует, то не понятно что будет
            {
                string sectionElId = null;

                File.Copy(Path.Combine(ModulesManager.GetModuleDirectory(moduleInfo.ShortName), sectionInfo.Name), Path.Combine(dictionarySectionPath, sectionInfo.Name), false);

                OneNoteUtils.UseOneNoteAPI(ref oneNoteApp, (oneNoteAppSafe) =>
                {
                    oneNoteAppSafe.OpenHierarchy(sectionInfo.Name, dictionarySectionId, out sectionElId, CreateFileType.cftSection);
                });

                if (string.IsNullOrEmpty(dictionarySectionId))
                    dictionarySectionId = sectionElId;
            }

            OneNoteUtils.UseOneNoteAPI(ref oneNoteApp, (oneNoteAppSafe) =>
            {
                oneNoteAppSafe.SyncHierarchy(dictionarySectionId);
            });

            return dictionarySectionId;
        }

        public static void WaitWhileDictionaryIsCreating(ref Application oneNoteApp, string dictionarySectionId, int? dictionaryPagesCount, int attemptsCount, Func<bool> checkIfExternalProcessAborted)
        {
            if (dictionaryPagesCount.HasValue)
            {
                if (attemptsCount < 500)  // 25 минут
                {
                    XmlNamespaceManager xnm;
                    var xDoc = OneNoteUtils.GetHierarchyElement(ref oneNoteApp, dictionarySectionId, HierarchyScope.hsPages, out xnm);
                    int pagesCount = xDoc.Root.XPathSelectElements("//one:Page", xnm).Count();
                    if (pagesCount < dictionaryPagesCount)
                    {
                        OneNoteUtils.UseOneNoteAPI(ref oneNoteApp, (oneNoteAppSafe) =>
                        {
                            oneNoteAppSafe.SyncHierarchy(dictionarySectionId);
                        });

                        for (var i = 0; i < 3; i++)
                        {
                            Thread.Sleep(1000);
                            if (checkIfExternalProcessAborted != null)
                            {
                                if (checkIfExternalProcessAborted())
                                    throw new ProcessAbortedByUserException();
                            }
                            System.Windows.Forms.Application.DoEvents();
                        }

                        WaitWhileDictionaryIsCreating(ref oneNoteApp, dictionarySectionId, dictionaryPagesCount, attemptsCount + 1, checkIfExternalProcessAborted);
                    }
                }
            }
        }

        public static void RemoveDictionary(ref Application oneNoteApp, string moduleName)
        {
            if (SettingsManager.Instance.DictionariesModules.Count == 1 
                && SettingsManager.Instance.DictionariesModules.First().ModuleName == moduleName)
            {
                CloseDictionariesNotebook(ref oneNoteApp);
            }
            else
            {
                var dictionaryModuleInfo = SettingsManager.Instance.DictionariesModules.FirstOrDefault(m => m.ModuleName == moduleName);
                if (dictionaryModuleInfo != null)
                {
                    try
                    {
                        OneNoteUtils.UseOneNoteAPI(ref oneNoteApp, (oneNoteAppSafe) =>
                        {
                            oneNoteAppSafe.DeleteHierarchy(dictionaryModuleInfo.SectionId);
                        });
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

        public static void CloseDictionariesNotebook(ref Application oneNoteApp)
        {
            OneNoteUtils.CloseNotebookSafe(ref oneNoteApp, SettingsManager.Instance.NotebookId_Dictionaries);

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
        public static bool GoToTerm(ref Application oneNoteApp, DictionaryTermLink link)
        {
            try
            {
                OneNoteUtils.UseOneNoteAPI(ref oneNoteApp, (oneNoteAppSafe) =>
                {
                    oneNoteAppSafe.NavigateTo(link.PageId, link.ObjectId);
                });
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
