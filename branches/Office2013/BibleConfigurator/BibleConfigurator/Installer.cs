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
            this.BeforeInstall += new InstallEventHandler(Installer_BeforeInstall);
            // Attach the 'Committed' event.
            this.Committed += new InstallEventHandler(MyInstaller_Committed);
            // Attach the 'Committing' event.
            this.Committing += new InstallEventHandler(MyInstaller_Committing);
        }

        void Installer_BeforeInstall(object sender, InstallEventArgs e)
        {
            try
            {
                string oneNoteTemplatesFolder = Path.Combine(Path.GetDirectoryName(Path.GetDirectoryName(Utils.GetCurrentDirectory())), "OneNoteTemplates");
                if ((!Directory.Exists(ModulesManager.GetModulesPackagesDirectory()) || Directory.GetFiles(ModulesManager.GetModulesPackagesDirectory()).Length == 0)
                    && Directory.Exists(oneNoteTemplatesFolder))
                {
                    if (!string.IsNullOrEmpty(SettingsManager.Instance.NotebookId_Bible))
                    {
                        SettingsManager.Instance.ModuleName = GenerateDefaultModule(oneNoteTemplatesFolder);
                        SettingsManager.Instance.Save();
                    }
                }
            }
            catch
            {
                //todo: log it
            }
        }

        // Event handler for 'Committing' event.
        private void MyInstaller_Committing(object sender, InstallEventArgs e)
        {            
            //Console.WriteLine("Committing Event occurred.");
            
        }

        // Event handler for 'Committed' event.
        private void MyInstaller_Committed(object sender, InstallEventArgs e)
        {
         
        }

        // Override the 'Install' method.
        public override void Install(IDictionary savedState)
        {
            base.Install(savedState);
        }

        // Override the 'Commit' method.
        public override void Commit(IDictionary savedState)
        {
            base.Commit(savedState);
        }

        // Override the 'Rollback' method.
        public override void Rollback(IDictionary savedState)
        {
            base.Rollback(savedState);
        }

        private static string GenerateDefaultModule(string oneNoteTemplatesFolder)
        {
            string defaultModuleName = "RST";
            string moduleFileName = defaultModuleName + Constants.IsbtFileExtension;
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
