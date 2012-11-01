using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;
using BibleCommon.Common;
using System.IO;
using BibleCommon.Consts;
using BibleCommon.Helpers;
using System.Threading;
using BibleCommon.Scheme;

namespace BibleCommon.Services
{
    public static class ModulesManager
    {
        private static Dictionary<Type, XmlSerializer> _serializers;        

        static ModulesManager()
        {
            _serializers = new Dictionary<Type, XmlSerializer>();
            _serializers.Add(typeof(ModuleInfo), new XmlSerializer(typeof(ModuleInfo)));
            _serializers.Add(typeof(XMLBIBLE), new XmlSerializer(typeof(XMLBIBLE)));
            _serializers.Add(typeof(ModuleDictionaryInfo), new XmlSerializer(typeof(ModuleDictionaryInfo)));
        }

        public static ModuleInfo GetCurrentModuleInfo()
        {
            if (!string.IsNullOrEmpty(SettingsManager.Instance.ModuleName))
                return GetModuleInfo(SettingsManager.Instance.ModuleName);

            throw new InvalidModuleException(BibleCommon.Resources.Constants.CurrentModuleIsUndefined);
        }

        public static string GetCurrentModuleDirectiory()
        {
            return GetModuleDirectory(SettingsManager.Instance.ModuleName);
        }

        public static string GetModuleDirectory(string moduleShortName)
        {
            return Path.Combine(GetModulesDirectory(), moduleShortName);
        }

        public static ModuleInfo GetModuleInfo(string moduleShortName)
        {   
            var module = GetModuleFile<ModuleInfo>(moduleShortName, Consts.Constants.ManifestFileName);
            if (string.IsNullOrEmpty(module.ShortName))
                module.ShortName = moduleShortName;
            
            return module;
        }

        public static int GetBibleChaptersCount(string moduleShortName, bool addBooksCount)
        {
            XMLBIBLE bibleInfo = null;
            int result;
            try
            {
                bibleInfo = GetModuleBibleInfo(moduleShortName);
                result = bibleInfo.Books.Sum(b => b.Chapters.Count);
                if (addBooksCount)
                    result += bibleInfo.Books.Count;
            }
            catch (InvalidModuleException)
            {
                result = 1189;
            }

            return result;
        }

        public static ModuleDictionaryInfo GetModuleDictionaryInfo(string moduleShortName)
        {
            return GetModuleFile<ModuleDictionaryInfo>(moduleShortName, Consts.Constants.DictionaryInfoFileName);
        }

        public static XMLBIBLE GetModuleBibleInfo(string moduleShortName)
        {
            return GetModuleFile<XMLBIBLE>(moduleShortName, Consts.Constants.BibleInfoFileName);
        }

        private static string GetModuleFilePath(string moduleShortName, string fileRelativePath)
        {
            string moduleDirectory = GetModuleDirectory(moduleShortName);
            string filePath = Path.Combine(moduleDirectory, fileRelativePath);
            if (!File.Exists(filePath))
                throw new InvalidModuleException(string.Format(BibleCommon.Resources.Constants.FileNotFound, filePath));

            return filePath;
        }

        private static T GetModuleFile<T>(string moduleShortName, string fileRelativePath)
        {
            var filePath = GetModuleFilePath(moduleShortName, fileRelativePath);

            return Dessirialize<T>(filePath);
        }

        public static void UpdateModuleManifest(ModuleInfo moduleInfo)
        {
            var filePath = GetModuleFilePath(moduleInfo.ShortName, Consts.Constants.ManifestFileName);

            Utils.SaveToXmlFile(moduleInfo, filePath);
        }

        private static T Dessirialize<T>(string xmlFilePath)
        {
            using (var fs = new FileStream(xmlFilePath, FileMode.Open))
            {
                return ((T)_serializers[typeof(T)].Deserialize(fs));
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

        public static List<ModuleInfo> GetModules(bool correctOnly)
        {
            var result = new List<ModuleInfo>();

            foreach (string moduleName in Directory.GetDirectories(GetModulesDirectory(), "*", SearchOption.TopDirectoryOnly))
            {
                try
                {
                    result.Add(GetModuleInfo(Path.GetFileName(moduleName)));
                }
                catch
                {
                    if (!correctOnly)
                        throw;
                }
            }

            return result;
        }        

        public static bool ModuleIsCorrect(string moduleName, Common.ModuleType? moduleType = null)
        {
            try
            {
                ModulesManager.CheckModule(moduleName, moduleType);
            }
            catch (InvalidModuleException)
            {              
                return false;
            }

            return true;
        }

        public static void CheckModule(string moduleDirectoryName, Common.ModuleType? moduleType = null)
        {
            ModuleInfo module = GetModuleInfo(moduleDirectoryName);
            
            CheckModule(module, moduleType);
        }

        public static void CheckModule(ModuleInfo module, Common.ModuleType? moduleType = null)
        {
            string moduleDirectory = GetModuleDirectory(module.ShortName);

            if (moduleType.HasValue)
                if (module.Type != moduleType.Value)
                    throw new InvalidModuleException(string.Format("Invalid module type: expected '{0}', actual '{1}'", moduleType, module.Type));

            if (module.MinProgramVersion != null)
            {
                var programVersion = Utils.GetProgramVersion();
                if (module.MinProgramVersion > programVersion)
                    throw new InvalidModuleException(string.Format(BibleCommon.Resources.Constants.ModuleIsNotSupported, module.MinProgramVersion, programVersion));
            }

            if (module.Type == Common.ModuleType.Bible)
            {
                var bibleModulePartTypes = new ContainerType[] { ContainerType.Bible, ContainerType.BibleStudy, ContainerType.BibleComments, ContainerType.BibleNotesPages };

                if (!module.UseSingleNotebook())
                {
                    foreach (var notebookType in bibleModulePartTypes)
                    {
                        if (!module.NotebooksStructure.Notebooks.Exists(n => n.Type == notebookType))
                            throw new InvalidModuleException(string.Format(Resources.Constants.Error_NotebookTemplateNotDefined, notebookType));
                    }
                }
                else
                {
                    foreach (var sectionGroupType in bibleModulePartTypes)
                    {
                        if (!module.GetNotebook(ContainerType.Single).SectionGroups.Exists(sg => sg.Type == sectionGroupType))
                            throw new InvalidModuleException(string.Format(Resources.Constants.Error_SectionGroupNotDefined, sectionGroupType, ContainerType.Single));
                    }
                }
            }

            foreach (var notebook in module.NotebooksStructure.Notebooks)
            {
                if (!notebook.SkipCheck)                
                    if (!File.Exists(Path.Combine(moduleDirectory, notebook.Name)))
                        throw new InvalidModuleException(string.Format(Resources.Constants.Error_NotebookTemplateNotFound, notebook.Name, notebook.Type));
            }

            foreach (var section in module.NotebooksStructure.Sections)
            {
                if (!section.SkipCheck)
                    if (!File.Exists(Path.Combine(moduleDirectory, section.Name)))
                        throw new InvalidModuleException(string.Format(Resources.Constants.Error_SectionFileNotFound, section.Name));
            }
        }

        public static ModuleInfo UploadModule(string originalFilePath, string destFilePath, string moduleName)
        {
            if (Path.GetExtension(originalFilePath).ToLower() != Constants.FileExtensionIsbt)
                throw new InvalidModuleException(string.Format(Resources.Constants.SelectFileWithExtension, Constants.FileExtensionIsbt)); 

            File.Copy(originalFilePath, destFilePath, true);

            string destFolder = GetModuleDirectory(moduleName);
            if (Directory.Exists(destFolder))
                Directory.Delete(destFolder, true);

            Directory.CreateDirectory(destFolder);

            try
            {
                ZipLibHelper.ExtractZipFile(File.ReadAllBytes(destFilePath), destFolder);
                var module = GetModuleInfo(moduleName);
                CheckModule(module);

                return module;
            }
            catch (Exception ex)
            {
                throw new InvalidModuleException(ex.Message);
            }
        }

        public static ModuleInfo ReadModuleInfo(string moduleFilePath)
        {            
            string destFolder = Path.Combine(Utils.GetTempFolderPath(), Path.GetFileNameWithoutExtension(moduleFilePath));
            try
            {
                if (Directory.Exists(destFolder))
                    Directory.Delete(destFolder, true);

                Directory.CreateDirectory(destFolder);

                ZipLibHelper.ExtractZipFile(File.ReadAllBytes(moduleFilePath), destFolder, new string[] { Constants.ManifestFileName });

                string manifestFilePath = Path.Combine(destFolder, Constants.ManifestFileName);
                if (!File.Exists(manifestFilePath))
                    throw new InvalidModuleException(string.Format(BibleCommon.Resources.Constants.FileNotFound, manifestFilePath));

                var module = Dessirialize<ModuleInfo>(manifestFilePath);
                if (string.IsNullOrEmpty(module.ShortName))
                    module.ShortName = Path.GetFileNameWithoutExtension(moduleFilePath);                

                return module;
            }
            finally
            {
                new Thread(DeleteDirectory).Start(destFolder);
            }
        }


        private static void DeleteDirectory(object directoryPath)
        {
            Thread.Sleep(500);
            try
            {
                if (Directory.Exists((string)directoryPath))
                    Directory.Delete((string)directoryPath, true);
            }
            catch { }                
        }

        public static void DeleteModule(string moduleShortName)
        {
            string moduleDirectory = GetModuleDirectory(moduleShortName);
            if (Directory.Exists(moduleDirectory))
                Directory.Delete(moduleDirectory, true);

            string manifestFilePath = Path.Combine(GetModulesPackagesDirectory(), moduleShortName + Constants.FileExtensionIsbt);
            if (File.Exists(manifestFilePath))
                File.Delete(manifestFilePath);
        }
    }
}
