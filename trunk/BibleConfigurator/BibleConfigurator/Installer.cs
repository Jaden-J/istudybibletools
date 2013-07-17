using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Configuration.Install;
using System.IO;
using System.Diagnostics;
using System.Reflection;
using System.Collections;
using BibleCommon.Helpers;
using BibleConfigurator.ModuleConverter;
using BibleCommon.Consts;
using BibleCommon.Services;
using System.Threading;
using BibleCommon.Common;
using Microsoft.Office.Interop.OneNote;

namespace BibleConfigurator
{

    // Taken from:http://msdn2.microsoft.com/en-us/library/
    // system.configuration.configurationmanager.aspx
    // Set 'RunInstaller' attribute to true.

    [RunInstaller(true)]
    public class Installer : System.Configuration.Install.Installer
    {
        private Version _programVersion;
        public Installer()
            : base()
        {
            try
            {   
                _programVersion = SettingsManager.Instance.VersionFromSettings;                

                // Attach the 'Committed' event.
                this.Committed += new InstallEventHandler(MyInstaller_Committed);                
            }
            catch (Exception ex)
            {
                Log(ex.ToString());
            }
        }

        // Event handler for 'Committed' event.
        private void MyInstaller_Committed(object sender, InstallEventArgs e)
        {
            TryToGenerateDefaultModule();            

            TryToRegenerateNotesPages();
            
            try
            {
                SettingsManager.Instance.Save();  // чтобы записать VersionFromSettings
            }
            catch (Exception ex)
            {
                Log(ex.ToString());
            }
        }

        public void TryToRegenerateNotesPages()
        {
            try
            {   
                if (_programVersion == null)  // так как до версии 3.1 этого свойства ещё не было
                {
                    var oneNoteApp = OneNoteUtils.CreateOneNoteAppSafe();                    
                    try
                    {
                        if (!string.IsNullOrEmpty(SettingsManager.Instance.FolderPath_BibleNotesPages))
                        {
                            var service = new AnalyzedVersesService(true);

                            AddDefaultAnalyzedNotebooksInfo(ref oneNoteApp, service);
                            RegenerateNotesPages(ref oneNoteApp, service);

                            service.Update();
                            NotesPageManagerFS.UpdateResources();
                        }
                    }
                    finally
                    {
                        OneNoteUtils.ReleaseOneNoteApp(ref oneNoteApp);                                                
                    }                    

                    TryToMergeAllModulesWithMainBible();  // так как раньше оно не правильно вызывалось и вызывалось ли вообще...
                }
            }
            catch (Exception ex)
            {
                Log(ex.ToString());
            }
        }

        private void RegenerateNotesPages(ref Application oneNoteApp, AnalyzedVersesService service)
        {
            if (Directory.Exists(SettingsManager.Instance.FolderPath_BibleNotesPages))
            {
                foreach (var filePath in Directory.GetFiles(SettingsManager.Instance.FolderPath_BibleNotesPages, "*.htm", SearchOption.AllDirectories))
                {
                    try
                    {
                        var fileContent = File.ReadAllText(filePath);
                        var startTitleIndex = fileContent.IndexOf("<title>") + "<title>".Length;
                        if (startTitleIndex > 10)
                        {
                            var endTitleIndex = fileContent.IndexOf("</title>");
                            var title = fileContent.Substring(startTitleIndex, endTitleIndex - startTitleIndex);
                            var parts = title.Split(new char[] { '[', ']' }, StringSplitOptions.RemoveEmptyEntries);
                            var pageName = parts[0];
                            var chapterPointer = new VersePointer(parts[1]);
                            var pageData = new NotesPageData(filePath, pageName, Path.GetFileNameWithoutExtension(filePath) == "0" ? NotesPageType.Chapter : NotesPageType.Verse, chapterPointer, true);
                            pageData.Serialize(ref oneNoteApp, service);
                        }

                        System.Windows.Forms.Application.DoEvents();
                    }
                    catch (Exception ex)
                    {
                        Log(ex.ToString());
                    }
                }
            }
        }

        private void AddDefaultAnalyzedNotebooksInfo(ref Application oneNoteApp, AnalyzedVersesService service)
        {
            foreach (var notebookInfo in SettingsManager.Instance.SelectedNotebooksForAnalyze)
            {
                try
                {
                    var notebookName = OneNoteUtils.GetHierarchyElementName(ref oneNoteApp, notebookInfo.NotebookId);
                    var notebookNickname = OneNoteUtils.GetNotebookElementNickname(ref oneNoteApp, notebookInfo.NotebookId);

                    service.AddAnalyzedNotebook(notebookName, notebookNickname);
                }
                catch (Exception ex)
                {
                    Log(ex.ToString());
                }               
            }
        }

        private void TryToMergeAllModulesWithMainBible()
        {
            try
            {
                BibleParallelTranslationManager.MergeAllModulesWithMainBible();
            }
            catch (Exception ex)
            {
                Log(ex.ToString());
            }
        }        

        private void TryToGenerateDefaultModule()
        {
            try
            {
                string oneNoteTemplatesFolder = Path.Combine(Path.GetDirectoryName(Path.GetDirectoryName(Utils.GetCurrentDirectory())), "OneNoteTemplates");
                if ((!Directory.Exists(ModulesManager.GetModulesPackagesDirectory()) || Directory.GetFiles(ModulesManager.GetModulesPackagesDirectory()).Length == 0)
                    && Directory.Exists(oneNoteTemplatesFolder))
                {
                    if (!string.IsNullOrEmpty(SettingsManager.Instance.NotebookId_Bible))
                    {
                        SettingsManager.Instance.ModuleShortName = GenerateDefaultModule(oneNoteTemplatesFolder);
                        SettingsManager.Instance.Save();
                    }
                }
            }
            catch (Exception ex)
            {
                Log(ex.ToString());
            }
        }

        private void Log(string message)
        {   
            try
            {
                var logFilePath = Path.Combine(Path.GetPathRoot(System.Environment.SystemDirectory), "isbt.log");
                File.AppendAllText(logFilePath, message + System.Environment.NewLine);

                if (Context != null)
                    Context.LogMessage(message);
            }
            catch { }
        }

        private static string GenerateDefaultModule(string oneNoteTemplatesFolder)
        {
            string defaultModuleName = "RST";
            string moduleFileName = defaultModuleName + Constants.FileExtensionIsbt;
            string tempFolderPath = Path.Combine(Utils.GetTempFolderPath(), "ModuleGenerator");
            if (!Directory.Exists(tempFolderPath))
                Directory.CreateDirectory(tempFolderPath);
            string destFilePath = Path.Combine(ModulesManager.GetModulesPackagesDirectory(), moduleFileName);

            DefaultRusModuleGenerator.GenerateModuleInfo(Path.Combine(tempFolderPath, Constants.ManifestFileName), SettingsManager.Instance.IsSingleNotebook);

            foreach (var filePath in Directory.GetFiles(oneNoteTemplatesFolder))
            {
                string fileName = Path.GetFileName(filePath);
                if (!SettingsManager.Instance.IsSingleNotebook)
                    if (fileName == "Holy Bible.onepkg")
                        continue;

                File.Copy(filePath, Path.Combine(tempFolderPath, fileName), true);
            }

            string isbtFilePath = Path.Combine(tempFolderPath, moduleFileName);

            ZipLibHelper.PackfilesToZip(tempFolderPath, isbtFilePath);

            ModulesManager.UploadModule(isbtFilePath, destFilePath, defaultModuleName);

            Thread.Sleep(500);

            try
            {
                Directory.Delete(tempFolderPath, true);
            }
            catch
            { }

            try
            {
                Directory.Delete(oneNoteTemplatesFolder, true);
            }
            catch
            { }

            return defaultModuleName;
        }
    }
}
