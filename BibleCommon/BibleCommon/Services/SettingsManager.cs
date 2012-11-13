﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Xml.Linq;
using System.Xml.XPath;
using BibleCommon.Helpers;
using Microsoft.Office.Interop.OneNote;
using System.Reflection;
using System.Threading;
using BibleCommon.Common;
using System.Resources;
using System.Globalization;
using System.ComponentModel;

namespace BibleCommon.Services
{
    public class StoredModuleInfo
    {
        public string ModuleName { get; set; }
        public Version ModuleVersion { get; set; }
        
        /// <summary>
        /// Section or SectionGroup ID
        /// </summary>
        public string SectionId { get; set; }

        public StoredModuleInfo(string moduleName, Version moduleVersion, string sectionId)
        {
            this.ModuleName = moduleName;
            this.ModuleVersion = moduleVersion;
            this.SectionId = sectionId;
        }

        public StoredModuleInfo(string moduleName, Version moduleVersion)
            :this(moduleName, moduleVersion, null)
        {
            
        }

        public StoredModuleInfo(string xmlString)
        {
            var parts = xmlString.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length < 2)
                throw new NotSupportedException(string.Format("Invalid StoredModuleInfo: '{0}'", xmlString));
            this.ModuleName = parts[0];
            this.ModuleVersion = new Version(parts[1]);
            if (parts.Length > 2)
                this.SectionId = parts[2];
        }

        public override string ToString()
        {
            if (this.SectionId != null)
                return string.Join(",", new string[] { this.ModuleName, this.ModuleVersion.ToString(), this.SectionId });
            else
                return string.Join(",", new string[] { this.ModuleName, this.ModuleVersion.ToString() });
        }
    }

    public class SettingsManager
    {
        #region Properties

        private static object _locker = new object();
        private string _filePath;
        private bool _useDefaultSettingsNodeExists = true;

        private static volatile SettingsManager _instance = null;
        public static SettingsManager Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (_locker)
                    {
                        if (_instance == null)
                        {
                            _instance = new SettingsManager();
                        }
                    }
                }

                return _instance;
            }
        }
        
        public string NotebookId_Bible { get; set; }
        public string NotebookId_BibleComments { get; set; }
        public string NotebookId_BibleNotesPages { get; set; }        
        public string NotebookId_BibleStudy { get; set; }
        public string NotebookId_SupplementalBible { get; set; }
        public string NotebookId_Dictionaries { get; set; }        
        public string SectionGroupId_Bible { get; set; }
        public string SectionGroupId_BibleStudy { get; set; }
        public string SectionGroupId_BibleComments { get; set; }        
        public string SectionGroupId_BibleNotesPages { get; set; }
        public string PageName_DefaultComments { get; set; }
        public string SectionName_DefaultBookOverview { get; set; }
        public string PageName_Notes { get; set; }
        public List<string> SelectedNotebooksForAnalyze { get; set; }

        private string _moduleName { get; set; }
        public string ModuleShortName 
        {
            get
            {
                if (!string.IsNullOrEmpty(_moduleName))
                    return _moduleName.ToLower();

                return _moduleName;
            }
            set
            {
                _moduleName = value;
            }
        }

        public Version NewVersionOnServer { get; set; }
        public DateTime? NewVersionOnServerLatestCheckTime { get; set; }
        public int PageWidth_Notes { get; set; }
        public int Language { get; set; }
        public int PageWidth_Bible { get; set; }
        public List<StoredModuleInfo> SupplementalBibleModules { get; set; }
        public string SupplementalBibleLinkName { get; set; }        
        public List<StoredModuleInfo> DictionariesModules { get; set; }

        /// <summary>
        /// Использовать ли промежуточные ссылки, которые открываются комманд-хэндлером OpenVerseHandler
        /// </summary>
        public bool UseMiddleStrongLinks { get; set; }
        

        /// <summary>
        /// Необходимо ли линковать каждый стих, входящий в MultiVerse
        /// </summary>
        public bool ExpandMultiVersesLinking { get; set; }

        /// <summary>
        /// необходимо ли линковать даже стихи, входящие в главу, помеченную в заголовке с []
        /// </summary>
        public bool ExcludedVersesLinking { get; set; }

        public bool UseDifferentPagesForEachVerse { get; set; }

        /// <summary>
        /// Параметры "мусорной" страницы
        /// </summary>
        public bool RubbishPage_Use { get; set; }
        public string PageName_RubbishNotes { get; set; }
        public int PageWidth_RubbishNotes { get; set; }
        public bool RubbishPage_ExpandMultiVersesLinking { get; set; }
        public bool RubbishPage_ExcludedVersesLinking { get; set; }

        public bool? UseDefaultSettings { get; set; }

        private Version _currentVersion = null;
        public Version CurrentVersion
        {
            get
            {
                if (_currentVersion == null)
                {
                    Assembly assembly = Assembly.GetCallingAssembly();
                    _currentVersion = assembly.GetName().Version;
                }

                return _currentVersion;
            }
        }



        /// <summary>
        /// Значение данного свойства сохраняется в памяти и не обновляется! нельзя использовать в коде, где текущий модуль может измениться
        /// </summary>
        private ModuleInfo _currentModule;
        public ModuleInfo CurrentModule
        {
            get
            {
                if (_currentModule == null)
                    _currentModule = ModulesManager.GetCurrentModuleInfo();

                return _currentModule;
            }
        }

        #endregion

        public string GetValidSupplementalBibleNotebookId(Application oneNoteApp, bool refreshCache = false)
        {
            if (!string.IsNullOrEmpty(NotebookId_SupplementalBible))
                if (!OneNoteUtils.NotebookExists(oneNoteApp, NotebookId_SupplementalBible, refreshCache))
                {
                    this.SupplementalBibleModules.Clear();
                    return null;
                }

            return NotebookId_SupplementalBible;
        }

        public string GetValidDictionariesNotebookId(Application oneNoteApp, bool refreshCache = false)
        {
            if (!string.IsNullOrEmpty(NotebookId_Dictionaries))
                if (!OneNoteUtils.NotebookExists(oneNoteApp, NotebookId_Dictionaries, refreshCache))
                {
                    this.DictionariesModules.Clear();
                    return null;
                }

            return NotebookId_Dictionaries;
        }

        public bool CurrentModuleIsCorrect()
        {
            return !string.IsNullOrEmpty(ModuleShortName) && ModulesManager.ModuleIsCorrect(ModuleShortName, ModuleType.Bible);
        }

        public bool IsConfigured(Application oneNoteApp)
        {
            bool result = !string.IsNullOrEmpty(this.NotebookId_Bible)
                && !string.IsNullOrEmpty(this.NotebookId_BibleComments)
                && !string.IsNullOrEmpty(this.NotebookId_BibleNotesPages)
                && !string.IsNullOrEmpty(this.NotebookId_BibleStudy)
                && !string.IsNullOrEmpty(this.SectionName_DefaultBookOverview)
                && !string.IsNullOrEmpty(this.PageName_DefaultComments)
                && !string.IsNullOrEmpty(this.PageName_Notes)
                && !string.IsNullOrEmpty(this.ModuleShortName)
                && ModulesManager.ModuleIsCorrect(this.ModuleShortName, ModuleType.Bible)
                && _useDefaultSettingsNodeExists;

            if (result)
            {
                if (this.IsSingleNotebook)
                {
                    result = !string.IsNullOrEmpty(this.SectionGroupId_Bible)
                          && !string.IsNullOrEmpty(this.SectionGroupId_BibleComments)
                          && !string.IsNullOrEmpty(this.SectionGroupId_BibleStudy)
                          && !string.IsNullOrEmpty(this.SectionGroupId_BibleNotesPages);                    

                    if (result)
                    {
                        result = OneNoteUtils.RootSectionGroupExists(oneNoteApp, this.NotebookId_Bible, this.SectionGroupId_Bible)
                            && OneNoteUtils.RootSectionGroupExists(oneNoteApp, this.NotebookId_Bible, this.SectionGroupId_BibleStudy)
                            && OneNoteUtils.RootSectionGroupExists(oneNoteApp, this.NotebookId_Bible, this.SectionGroupId_BibleComments)
                            && OneNoteUtils.RootSectionGroupExists(oneNoteApp, this.NotebookId_Bible, this.SectionGroupId_BibleNotesPages);
                    }
                }

                if (result)
                {
                    result = OneNoteUtils.NotebookExists(oneNoteApp, this.NotebookId_Bible)
                        && OneNoteUtils.NotebookExists(oneNoteApp, this.NotebookId_BibleComments)
                        && OneNoteUtils.NotebookExists(oneNoteApp, this.NotebookId_BibleStudy)
                        && OneNoteUtils.NotebookExists(oneNoteApp, this.NotebookId_BibleNotesPages);
                }
            }

            return result;
        }        

        public bool IsSingleNotebook
        {
            get
            {
                return this.NotebookId_Bible == this.NotebookId_BibleComments
                    && this.NotebookId_Bible == this.NotebookId_BibleStudy
                    && this.NotebookId_Bible == this.NotebookId_BibleNotesPages;
            }
        }        

        public void ReLoadSettings()
        {
            if (!File.Exists(_filePath))
                LoadDefaultSettings();
            else
                LoadSettingsFromFile();        
        }

        protected SettingsManager()
        {
            SetDefaultSettings();

            _filePath = Path.Combine(Utils.GetProgramDirectory(), Consts.Constants.ConfigFileName);

            ReLoadSettings();
        }

        private void SetDefaultSettings()
        {
            SupplementalBibleModules = new List<StoredModuleInfo>();
            DictionariesModules = new List<StoredModuleInfo>();
        }

        private void LoadSettingsFromFile()
        {
            XDocument xdoc = XDocument.Load(_filePath);

            try
            {   
                LoadGeneralSettings(xdoc);
                LoadAdditionalSettings(xdoc);

                this.UseDefaultSettings = GetParameterValue<bool?>(xdoc, Consts.Constants.ParameterName_UseDefaultSettings, null, s => bool.Parse(s));

                bool programSettingsWasLoaded = false;
                if (!UseDefaultSettings.HasValue)
                {
                    LoadProgramSettings(xdoc);
                    programSettingsWasLoaded = true;
                    _useDefaultSettingsNodeExists = false;

                    this.UseDefaultSettings = DetermineIfCurrentSettingsAreDefualt();                    
                }

                if (UseDefaultSettings.Value)
                    LoadDefaultSettings();                
                else if (!programSettingsWasLoaded)
                    LoadProgramSettings(xdoc);
            }
            catch (Exception ex)
            {
                Logger.LogError(ex.Message);
                LoadDefaultSettings();
            }
        }        

        private void LoadAdditionalSettings(XDocument xdoc)
        {
            this.NewVersionOnServer = GetParameterValue<Version>(xdoc, Consts.Constants.ParameterName_NewVersionOnServer, null, value => new Version(value));
            this.NewVersionOnServerLatestCheckTime = GetParameterValue<DateTime?>(xdoc, Consts.Constants.ParameterName_NewVersionOnServerLatestCheckTime, null, value => DateTime.Parse(value));            
            this.Language = GetParameterValue<int>(xdoc, Consts.Constants.ParameterName_Language, Thread.CurrentThread.CurrentUICulture.LCID);
            this.ModuleShortName = GetParameterValue<string>(xdoc, Consts.Constants.ParameterName_ModuleName, string.Empty);            
        }

        private CultureInfo _currentCultureInfo = null;
        public CultureInfo CurrentResourceCulture
        {
            get
            {
                if (_currentCultureInfo == null)
                {

                    if (this.Language == 0)
                        this.Language = Thread.CurrentThread.CurrentUICulture.LCID;

                    _currentCultureInfo = new CultureInfo(this.Language);       // потому что локаль текущего потока может быть ещё не установлена
                }

                return _currentCultureInfo;
            }
        }

        public string GetResourceString(string resourceName)
        {
            return Resources.Constants.ResourceManager.GetString(resourceName, CurrentResourceCulture);
        }

        /// <summary>
        /// Эти настройки сбрасываются, если UseDefaultSettings == true
        /// </summary>
        private void LoadProgramSettings(XDocument xdoc)
        {
            this.SectionName_DefaultBookOverview = GetParameterValue<string>(xdoc, Consts.Constants.ParameterName_SectionNameDefaultBookOverview,
                                                        GetResourceString(Consts.Constants.ResourceName_DefaultPageNameDefaultBookOverview));
            this.PageName_DefaultComments = GetParameterValue<string>(xdoc, Consts.Constants.ParameterName_PageNameDefaultComments);
            this.PageName_Notes = GetParameterValue<string>(xdoc, Consts.Constants.ParameterName_PageNameNotes);
            this.PageWidth_Notes = GetParameterValue<int>(xdoc, Consts.Constants.ParameterName_PageWidthNotes, 500);
            this.PageWidth_Bible = GetParameterValue<int>(xdoc, Consts.Constants.ParameterName_PageWidthBible, 500);
            this.ExpandMultiVersesLinking = GetParameterValue<bool>(xdoc, Consts.Constants.ParameterName_ExpandMultiVersesLinking);
            this.ExcludedVersesLinking = GetParameterValue<bool>(xdoc, Consts.Constants.ParameterName_ExcludedVersesLinking);
            this.UseDifferentPagesForEachVerse = GetParameterValue<bool>(xdoc, Consts.Constants.ParameterName_UseDifferentPagesForEachVerse);
            this.RubbishPage_Use = GetParameterValue<bool>(xdoc, Consts.Constants.ParameterName_RubbishPageUse);
            this.PageName_RubbishNotes = GetParameterValue<string>(xdoc, Consts.Constants.ParameterName_PageNameRubbishNotes,
                                                        GetResourceString(Consts.Constants.ResourceName_DefaultPageName_RubbishNotes));
            this.PageWidth_RubbishNotes = GetParameterValue<int>(xdoc, Consts.Constants.ParameterName_PageWidthRubbishNotes, 500);
            this.RubbishPage_ExpandMultiVersesLinking = GetParameterValue<bool>(xdoc, Consts.Constants.ParameterName_RubbishPageExpandMultiVersesLinking, Consts.Constants.DefaultRubbishPage_ExpandMultiVersesLinking);
            this.RubbishPage_ExcludedVersesLinking = GetParameterValue<bool>(xdoc, Consts.Constants.ParameterName_RubbishPageExcludedVersesLinking, Consts.Constants.DefaultRubbishPage_ExcludedVersesLinking);
            this.UseMiddleStrongLinks = GetParameterValue<bool>(xdoc, Consts.Constants.ParameterName_UseMiddleStrongLinks, Consts.Constants.DefaultUseMiddleStrongLinks);            
        }

        private void LoadGeneralSettings(XDocument xdoc)
        {
            this.NotebookId_Bible = GetParameterValue<string>(xdoc, Consts.Constants.ParameterName_NotebookIdBible);
            this.NotebookId_BibleComments = GetParameterValue<string>(xdoc, Consts.Constants.ParameterName_NotebookIdBibleComments);
            this.NotebookId_BibleNotesPages = GetParameterValue<string>(xdoc, Consts.Constants.ParameterName_NotebookIdBibleNotesPages);
            this.NotebookId_BibleStudy = GetParameterValue<string>(xdoc, Consts.Constants.ParameterName_NotebookIdBibleStudy);            
            this.SectionGroupId_Bible = GetParameterValue<string>(xdoc, Consts.Constants.ParameterName_SectionGroupIdBible);
            this.SectionGroupId_BibleStudy = GetParameterValue<string>(xdoc, Consts.Constants.ParameterName_SectionGroupIdBibleStudy);
            this.SectionGroupId_BibleComments = GetParameterValue<string>(xdoc, Consts.Constants.ParameterName_SectionGroupIdBibleComments);
            this.SectionGroupId_BibleNotesPages = GetParameterValue<string>(xdoc, Consts.Constants.ParameterName_SectionGroupIdBibleNotesPages);

            this.NotebookId_SupplementalBible = GetParameterValue<string>(xdoc, Consts.Constants.ParameterName_NotebookIdSupplementalBible);
            this.SupplementalBibleModules = GetParameterValue<List<StoredModuleInfo>>(xdoc, Consts.Constants.ParameterName_SupplementalBibleModules, new List<StoredModuleInfo>(),
                                                s => s.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries).ToList().ConvertAll(xmlString => new StoredModuleInfo(xmlString)));
            this.SupplementalBibleLinkName = GetParameterValue<string>(xdoc, Consts.Constants.ParameterName_SupplementalBibleLinkName,
                                                  GetResourceString(Consts.Constants.ResourceName_DefaultSupplementalBibleLinkName));

            this.NotebookId_Dictionaries = GetParameterValue<string>(xdoc, Consts.Constants.ParameterName_NotebookIdDictionaries);
            this.DictionariesModules = GetParameterValue<List<StoredModuleInfo>>(xdoc, Consts.Constants.ParameterName_DictionariesModules, new List<StoredModuleInfo>(), 
                                                s => s.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries).ToList().ConvertAll(xmlString => new StoredModuleInfo(xmlString)));

            this.SelectedNotebooksForAnalyze = GetParameterValue<List<string>>(xdoc, Consts.Constants.ParameterName_SelectedNotebooksForAnalyze, 
                                                        GetDefaultNotebooksForAnalyze(), GetNotebooksForAnalyze); 
        }

        private T GetParameterValue<T>(XDocument xdoc, string parameterName, object defaultValue = null, Func<string, T> convertFunc = null)
        {
            XElement el = xdoc.Root.XPathSelectElement(parameterName);

            if (el == null || string.IsNullOrEmpty(el.Value))
            {
                if (defaultValue != null)
                    return (T)defaultValue;
                else
                    return default(T);                
            }
            else
            {
                if (convertFunc != null)
                    return convertFunc(el.Value);
                else
                    return ConvertFromString<T>(el.Value);
            }
        }

        private static T ConvertFromString<T>(string value)
        {
            if (string.IsNullOrEmpty(value))
                return default(T);

            Type typeParameterType = typeof(T);

            return (T)Convert.ChangeType(value, typeParameterType);
        }


        /// <summary>
        /// Загружает настройки по умолчанию, если стоит галочка "Использовать настройки по умолчанию"
        /// </summary>
        public void LoadDefaultSettings()
        {
            this.UseDefaultSettings = true;      
            
            this.PageWidth_Notes = Consts.Constants.DefaultPageWidth_Notes;
            this.PageWidth_Bible = Consts.Constants.DefaultPageWidth_Bible;
            this.ExpandMultiVersesLinking = Consts.Constants.DefaultExpandMultiVersesLinking;            
            this.ExcludedVersesLinking = Consts.Constants.DefaultExcludedVersesLinking;
            this.UseDifferentPagesForEachVerse = Consts.Constants.DefaultUseDifferentPagesForEachVerse;
            this.RubbishPage_Use = Consts.Constants.DefaultRubbishPage_Use;            
            this.PageWidth_RubbishNotes = Consts.Constants.DefaultPageWidth_RubbishNotes;
            this.RubbishPage_ExpandMultiVersesLinking = Consts.Constants.DefaultRubbishPage_ExpandMultiVersesLinking;
            this.RubbishPage_ExcludedVersesLinking = Consts.Constants.DefaultRubbishPage_ExcludedVersesLinking;
            this.UseMiddleStrongLinks = Consts.Constants.DefaultUseMiddleStrongLinks;

            LoadDefaultLocalazibleSettings();
        }

        public void LoadDefaultLocalazibleSettings()
        {
            this.SectionName_DefaultBookOverview = GetResourceString(Consts.Constants.ResourceName_DefaultPageNameDefaultBookOverview);
            this.PageName_DefaultComments = GetResourceString(Consts.Constants.ResourceName_DefaultPageNameDefaultComments);
            this.PageName_Notes = GetResourceString(Consts.Constants.ResourceName_DefaultPageName_Notes);
            this.PageName_RubbishNotes = GetResourceString(Consts.Constants.ResourceName_DefaultPageName_RubbishNotes);
        }

        private bool DetermineIfCurrentSettingsAreDefualt()
        {
            return this.SectionName_DefaultBookOverview == GetResourceString(Consts.Constants.ResourceName_DefaultPageNameDefaultBookOverview)
                && this.PageName_DefaultComments == GetResourceString(Consts.Constants.ResourceName_DefaultPageNameDefaultComments)
                && this.PageName_Notes == GetResourceString(Consts.Constants.ResourceName_DefaultPageName_Notes)
                && this.PageName_RubbishNotes == GetResourceString(Consts.Constants.ResourceName_DefaultPageName_RubbishNotes)
                && this.PageWidth_Notes == Consts.Constants.DefaultPageWidth_Notes
                && this.ExpandMultiVersesLinking == false
                && this.ExcludedVersesLinking == Consts.Constants.DefaultExcludedVersesLinking
                && this.UseDifferentPagesForEachVerse == false
                && this.RubbishPage_Use == Consts.Constants.DefaultRubbishPage_Use                
                && this.PageWidth_RubbishNotes == Consts.Constants.DefaultPageWidth_RubbishNotes
                && this.RubbishPage_ExpandMultiVersesLinking == Consts.Constants.DefaultRubbishPage_ExpandMultiVersesLinking
                && this.RubbishPage_ExcludedVersesLinking == Consts.Constants.DefaultRubbishPage_ExcludedVersesLinking
                && this.UseMiddleStrongLinks == Consts.Constants.DefaultUseMiddleStrongLinks;
        }

        public void Save()
        {
            using (FileStream fs = new FileStream(_filePath, FileMode.Create))
            {
                using (StreamWriter sw = new StreamWriter(fs))
                {
                    XDocument xDoc = XDocument.Parse("<Settings></Settings>");

                    xDoc.Root.Add(new XElement(Consts.Constants.ParameterName_NotebookIdBible, this.NotebookId_Bible),
                                  new XElement(Consts.Constants.ParameterName_NotebookIdBibleComments, this.NotebookId_BibleComments),
                                  new XElement(Consts.Constants.ParameterName_NotebookIdBibleNotesPages, this.NotebookId_BibleNotesPages),
                                  new XElement(Consts.Constants.ParameterName_NotebookIdBibleStudy, this.NotebookId_BibleStudy),                                  
                                  new XElement(Consts.Constants.ParameterName_NotebookIdSupplementalBible, this.NotebookId_SupplementalBible),
                                  new XElement(Consts.Constants.ParameterName_NotebookIdDictionaries, this.NotebookId_Dictionaries),
                                  new XElement(Consts.Constants.ParameterName_SectionGroupIdBible, this.SectionGroupId_Bible),                                                                    
                                  new XElement(Consts.Constants.ParameterName_SectionGroupIdBibleComments, this.SectionGroupId_BibleComments),
                                  new XElement(Consts.Constants.ParameterName_SectionGroupIdBibleNotesPages, this.SectionGroupId_BibleNotesPages),
                                  new XElement(Consts.Constants.ParameterName_SectionGroupIdBibleStudy, this.SectionGroupId_BibleStudy),
                                  new XElement(Consts.Constants.ParameterName_SectionNameDefaultBookOverview, this.SectionName_DefaultBookOverview),
                                  new XElement(Consts.Constants.ParameterName_PageNameDefaultComments, this.PageName_DefaultComments),
                                  new XElement(Consts.Constants.ParameterName_PageNameNotes, this.PageName_Notes),                                  
                                  new XElement(Consts.Constants.ParameterName_NewVersionOnServer, this.NewVersionOnServer),
                                  new XElement(Consts.Constants.ParameterName_NewVersionOnServerLatestCheckTime, this.NewVersionOnServerLatestCheckTime.HasValue
                                                ? this.NewVersionOnServerLatestCheckTime.Value.ToString() : string.Empty),
                                  new XElement(Consts.Constants.ParameterName_PageWidthNotes, this.PageWidth_Notes),
                                  new XElement(Consts.Constants.ParameterName_ExpandMultiVersesLinking, this.ExpandMultiVersesLinking),
                                  new XElement(Consts.Constants.ParameterName_ExcludedVersesLinking, this.ExcludedVersesLinking),
                                  new XElement(Consts.Constants.ParameterName_UseDifferentPagesForEachVerse, this.UseDifferentPagesForEachVerse),
                                  new XElement(Consts.Constants.ParameterName_RubbishPageUse, this.RubbishPage_Use),
                                  new XElement(Consts.Constants.ParameterName_PageNameRubbishNotes, this.PageName_RubbishNotes),
                                  new XElement(Consts.Constants.ParameterName_PageWidthRubbishNotes, this.PageWidth_RubbishNotes),
                                  new XElement(Consts.Constants.ParameterName_RubbishPageExpandMultiVersesLinking, this.RubbishPage_ExpandMultiVersesLinking),
                                  new XElement(Consts.Constants.ParameterName_RubbishPageExcludedVersesLinking, this.RubbishPage_ExcludedVersesLinking),
                                  new XElement(Consts.Constants.ParameterName_Language, this.Language),
                                  new XElement(Consts.Constants.ParameterName_ModuleName, this.ModuleShortName),
                                  new XElement(Consts.Constants.ParameterName_UseDefaultSettings, this.UseDefaultSettings.Value),
                                  new XElement(Consts.Constants.ParameterName_PageWidthBible, this.PageWidth_Bible),
                                  new XElement(Consts.Constants.ParameterName_SupplementalBibleModules, string.Join(";", this.SupplementalBibleModules.ConvertAll(dm => dm.ToString()).ToArray())),
                                  new XElement(Consts.Constants.ParameterName_SupplementalBibleLinkName, this.SupplementalBibleLinkName),                                  
                                  new XElement(Consts.Constants.ParameterName_DictionariesModules, string.Join(";", this.DictionariesModules.ConvertAll(dm => dm.ToString()).ToArray())),
                                  new XElement(Consts.Constants.ParameterName_UseMiddleStrongLinks, UseMiddleStrongLinks)                                  
                                  );

                    if (SelectedNotebooksForAnalyze != GetDefaultNotebooksForAnalyze())
                        xDoc.Root.Add(new XElement(Consts.Constants.ParameterName_SelectedNotebooksForAnalyze, ConvertNotebooksForAnalyzeToString(SelectedNotebooksForAnalyze)));

                    xDoc.Save(sw);
                    sw.Flush();                    
                }
            }
        }

        private List<string> GetNotebooksForAnalyze(string selectedNotebooksString)
        {
            if (!string.IsNullOrEmpty(selectedNotebooksString))
            {
                string[] s = selectedNotebooksString.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);

                if (s.Length == 2 && s[0] == IsSingleNotebook.ToString().ToLower())
                {
                    return s[1].Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                }
            }

            return GetDefaultNotebooksForAnalyze();
        }

        private List<string> GetDefaultNotebooksForAnalyze()
        {
            if (IsSingleNotebook)
            {
                return new List<string>() 
                    {
                        SectionGroupId_BibleStudy, 
                        SectionGroupId_BibleComments
                    };
            }
            else
            {
                return new List<string>() 
                    {
                        NotebookId_BibleStudy, 
                        NotebookId_BibleComments
                    };
            }

        }

        private string ConvertNotebooksForAnalyzeToString(List<string> notebooksIds)
        {
            return string.Join(";", new string[] 
            {
                IsSingleNotebook.ToString().ToLower(), string.Join(",", notebooksIds.ToArray())
            });
        } 
    }
}


