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

        public static ModuleInfo GetCurrentModuleInfo()
        {
            return GetModuleInfo(SettingsManager.Instance.ModuleName);
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
                throw new InvalidModuleException(string.Format("File '{0}' was not found.", manifestFilePath));

            using (var fs = new FileStream(manifestFilePath, FileMode.Open))
            {
                return ((ModuleInfo)_serializer.Deserialize(fs));
            }
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

            foreach (NotebookType notebookType in Enum.GetValues(typeof(NotebookType)))
            {
                if (!module.Notebooks.Exists(n => n.Type == notebookType))
                    throw new Exception(string.Format("Не указан шаблон записной книжки типа '{0}'.", notebookType));  //todo: локализовать
            }

            foreach (SectionGroupType sectionGroupType in Enum.GetValues(typeof(SectionGroupType)))
            {
                if (!module.GetNotebook(NotebookType.Single).SectionGroups.Exists(sg => sg.Type == sectionGroupType))
                    throw new Exception(string.Format("Не указана группа разделов типа '{0}' в шаблоне записной книжки типа '{1}'.", sectionGroupType, NotebookType.Single));  //todo: локализовать
            }

            foreach (var notebook in module.Notebooks)
            {
                if (!File.Exists(Path.Combine(moduleDirectory, notebook.Name)))
                    throw new Exception(string.Format("Не найден шаблон записной книжки '{0}' типа '{1}'.", notebook.Name, notebook.Type));  //todo: локализовать
            }            
        }

        public static void UploadModule(string originalFilePath, string destFilePath, string moduleName)
        {
            if (Path.GetExtension(originalFilePath).ToLower() != Constants.IsbtFileExtension)
                throw new InvalidModuleException(string.Format("Выберите файл с расширением '{0}'", Constants.IsbtFileExtension));   //todo: локализовать

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
