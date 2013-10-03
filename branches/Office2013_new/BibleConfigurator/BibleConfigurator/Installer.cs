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
using BibleCommon.UI.Forms;

namespace BibleConfigurator
{

    // Taken from:http://msdn2.microsoft.com/en-us/library/
    // system.configuration.configurationmanager.aspx
    // Set 'RunInstaller' attribute to true.

    [RunInstaller(true)]
    public class Installer : System.Configuration.Install.Installer
    {   
        public Installer()
            : base()
        {
            DoWithExceptionHandling(() =>
            {
                this.Committed += new InstallEventHandler(MyInstaller_Committed);
            });
        }

        private void DoWithExceptionHandling(Action action)
        {
            try
            {
                if (action != null)
                    action();
            }
            catch (Exception ex)
            {
                Log(ex.ToString());
            }
        }

        // Event handler for 'Committed' event.
        private void MyInstaller_Committed(object sender, InstallEventArgs e)
        {
            DoWithExceptionHandling(TryToGenerateDefaultModule);
        }      

        private void TryToGenerateDefaultModule()
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

        private void Log(string message)
        {   
            try
            {
                var logFilePath = Path.Combine(Path.GetPathRoot(System.Environment.SystemDirectory), "isbt.log");
                File.AppendAllText(logFilePath, string.Format("{0}: {1}{2}", DateTime.Now, message, System.Environment.NewLine));

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
