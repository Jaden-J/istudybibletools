using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;
using BibleCommon.Common;
using System.IO;
using BibleCommon.Consts;

namespace BibleCommon.Services
{
    public static class ModulesManager
    {
        private static XmlSerializer _serializer;

        static ModulesManager()
        {
            _serializer = new XmlSerializer(typeof(ModuleInfo));
        }

        public static ModuleInfo GetModuleInfo(string moduleDirectoryName)
        {
            string moduleDirectory = Path.Combine(GetModulesDirectory(), moduleDirectoryName);
            string manifestFilePath = Path.Combine(moduleDirectory, Consts.Constants.ManifestFileName);
            if (File.Exists(manifestFilePath))
            {
                using (var fs = new FileStream(manifestFilePath, FileMode.Open))
                {
                    return ((ModuleInfo)_serializer.Deserialize(fs));
                }
            }

            return null;
        }

        public static string GetModulesDirectory()
        {
            string directoryPath = SettingsManager.GetProgramDirectory();

            string modulesDirectory = Path.Combine(directoryPath, Constants.ModulesDirectoryName);

            if (!Directory.Exists(modulesDirectory))
                Directory.CreateDirectory(modulesDirectory);

            return modulesDirectory;
        }

        public static string GetModulesPackagesDirectory()
        {
            string directoryPath = SettingsManager.GetProgramDirectory();

            string modulesDirectory = Path.Combine(directoryPath, Constants.ModulesPackagesDirectoryName);

            if (!Directory.Exists(modulesDirectory))
                Directory.CreateDirectory(modulesDirectory);

            return modulesDirectory;
        }

        private static void CheckModule(string moduleDirectoryName)
        {
            ModuleInfo module = GetModuleInfo(moduleDirectoryName);

            string moduleDirectory = Path.Combine(GetModulesDirectory(), moduleDirectoryName);

            foreach (var notebook in module.Notebooks)
            {
                if (!File.Exists(Path.Combine(moduleDirectory, notebook.Name)))
                    throw new Exception(string.Format("Не найдена записная книжка '{0}' типа '{1}'.", notebook.Name, notebook.Type));  //todo: локализовать
            }
        }

        public static void UploadModule(string originalFilePath, string destFilePath, string moduleName)
        {
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
            Directory.Delete(moduleDirectory, true);

            string manifestFilePath = Path.Combine(GetModulesPackagesDirectory(), moduleDirectoryName + Constants.IsbtFileExtension);
            File.Delete(manifestFilePath);
        }
    }
}
