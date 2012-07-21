using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;
using BibleCommon.Common;
using System.IO;
using BibleCommon.Consts;
using BibleCommon.Helpers;

namespace BibleCommon.Services
{
    public static class ModulesManager
    {
        private static XmlSerializer _serializer;

        static ModulesManager()
        {
            _serializer = new XmlSerializer(typeof(ModuleInfo));
        }

        public static ModuleInfo GetCurrentModuleInfo()
        {
            if (!string.IsNullOrEmpty(SettingsManager.Instance.ModuleName))
                return GetModuleInfo(SettingsManager.Instance.ModuleName);

            throw new InvalidModuleException(BibleCommon.Resources.Constants.CurrentModuleIsUndefined);
        }

        public static string GetCurrentModuleDirectiory()
        {
            return Path.Combine(GetModulesDirectory(), SettingsManager.Instance.ModuleName);
        }

        public static ModuleInfo GetModuleInfo(string moduleDirectoryName)
        {
            string moduleDirectory = Path.Combine(GetModulesDirectory(), moduleDirectoryName);
            string manifestFilePath = Path.Combine(moduleDirectory, Consts.Constants.ManifestFileName);
            if (!File.Exists(manifestFilePath))
                throw new InvalidModuleException(string.Format(BibleCommon.Resources.Constants.FileNotFound, manifestFilePath));

            using (var fs = new FileStream(manifestFilePath, FileMode.Open))
            {
                var module = ((ModuleInfo)_serializer.Deserialize(fs));
                module.ShortName = moduleDirectoryName;
                return module;
            }
        }       

        public static string GetModulesDirectory()
        {
            string directoryPath = Utils.GetProgramDirectory();

            string modulesDirectory = Path.Combine(directoryPath, Constants.ModulesDirectoryName);

            if (!Directory.Exists(modulesDirectory))
                Directory.CreateDirectory(modulesDirectory);

            return modulesDirectory;
        }

        public static string GetModulesPackagesDirectory()
        {
            string directoryPath = Utils.GetProgramDirectory();

            string modulesDirectory = Path.Combine(directoryPath, Constants.ModulesPackagesDirectoryName);

            if (!Directory.Exists(modulesDirectory))
                Directory.CreateDirectory(modulesDirectory);

            return modulesDirectory;
        }

        public static bool ModuleIsCorrect(string moduleName)
        {
            try
            {
                ModulesManager.CheckModule(moduleName);
            }
            catch
            {              
                return false;
            }

            return true;
        }

        public static void CheckModule(string moduleDirectoryName)
        {
            ModuleInfo module = GetModuleInfo(moduleDirectoryName);

            string moduleDirectory = Path.Combine(GetModulesDirectory(), moduleDirectoryName);

            foreach (NotebookType notebookType in Enum.GetValues(typeof(NotebookType)).Cast<NotebookType>().Where(t => t != NotebookType.Single))
            {
                if (!module.Notebooks.Exists(n => n.Type == notebookType))
                    throw new Exception(string.Format(Resources.Constants.Error_NotebookTemplateNotDefined, notebookType)); 
            }

            if (module.UseSingleNotebook())
            {
                foreach (SectionGroupType sectionGroupType in Enum.GetValues(typeof(SectionGroupType)))
                {
                    if (!module.GetNotebook(NotebookType.Single).SectionGroups.Exists(sg => sg.Type == sectionGroupType))
                        throw new Exception(string.Format(Resources.Constants.Error_SectionGroupNotDefined, sectionGroupType, NotebookType.Single));
                }
            }

            foreach (var notebook in module.Notebooks)
            {
                if (!File.Exists(Path.Combine(moduleDirectory, notebook.Name)))
                    throw new Exception(string.Format(Resources.Constants.Error_NotebookTemplateNotFound, notebook.Name, notebook.Type));  
            }            
        }

        public static void UploadModule(string originalFilePath, string destFilePath, string moduleName)
        {
            if (Path.GetExtension(originalFilePath).ToLower() != Constants.IsbtFileExtension)
                throw new InvalidModuleException(string.Format(Resources.Constants.SelectFileWithExtension, Constants.IsbtFileExtension)); 

            File.Copy(originalFilePath, destFilePath, true);

            string destFolder = Path.Combine(ModulesManager.GetModulesDirectory(), moduleName);
            if (Directory.Exists(destFolder))
                Directory.Delete(destFolder, true);

            Directory.CreateDirectory(destFolder);

            try
            {
                ZipLibHelper.ExtractZipFile(File.ReadAllBytes(destFilePath), destFolder);
                CheckModule(moduleName);
            }
            catch (Exception ex)
            {
                throw new InvalidModuleException(ex.Message);
            }
        }

        public static void DeleteModule(string moduleDirectoryName)
        {
            string moduleDirectory = Path.Combine(GetModulesDirectory(), moduleDirectoryName);
            if (Directory.Exists(moduleDirectory))
                Directory.Delete(moduleDirectory, true);

            string manifestFilePath = Path.Combine(GetModulesPackagesDirectory(), moduleDirectoryName + Constants.IsbtFileExtension);
            if (File.Exists(manifestFilePath))
                File.Delete(manifestFilePath);
        }
    }
}
