using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Xml.Linq;
using System.Xml.XPath;
using BibleCommon.Helpers;
using Microsoft.Office.Interop.OneNote;
using System.Reflection;

namespace BibleCommon.Services
{
    public class SettingsManager
    {
        private static object _locker = new object();
        private string _filePath;

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
        public string NotebookId_BibleStudy { get; set; }        
        public string SectionGroupId_Bible { get; set; }
        public string SectionGroupId_BibleComments { get; set; }
        public string SectionGroupId_BibleStudy { get; set; }
        public string PageName_DefaultComments { get; set; }
        public string PageName_DefaultBookOverview { get; set; }
        public string PageName_Notes { get; set; }

        public DateTime? LastNotesLinkTime { get; set; }
        public Version NewVersionOnServer { get; set; }
        public DateTime? NewVersionOnServerLatestCheckTime { get; set; }

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

        public bool IsConfigured(Application oneNoteApp)
        {
            bool result = !string.IsNullOrEmpty(this.NotebookId_Bible)
                && !string.IsNullOrEmpty(this.NotebookId_BibleComments)
                && !string.IsNullOrEmpty(this.NotebookId_BibleStudy)
                && !string.IsNullOrEmpty(this.PageName_DefaultBookOverview)
                && !string.IsNullOrEmpty(this.PageName_DefaultComments)
                && !string.IsNullOrEmpty(this.PageName_Notes);

            if (result)
            {
                if (this.IsSingleNotebook)
                {
                    result = !string.IsNullOrEmpty(this.SectionGroupId_Bible)
                          && !string.IsNullOrEmpty(this.SectionGroupId_BibleComments)
                          && !string.IsNullOrEmpty(this.SectionGroupId_BibleStudy);

                    //todo: чтоб проверял и наличие секций
                }
                
                result = OneNoteUtils.NotebookExists(oneNoteApp, this.NotebookId_Bible)
                    && OneNoteUtils.NotebookExists(oneNoteApp, this.NotebookId_BibleComments)
                    && OneNoteUtils.NotebookExists(oneNoteApp, this.NotebookId_BibleStudy);
            }

            return result;
        }

        public bool IsSingleNotebook
        {
            get
            {
                return this.NotebookId_Bible == this.NotebookId_BibleComments
                    && this.NotebookId_Bible == this.NotebookId_BibleStudy;
            }
        }       

        protected SettingsManager()
        {
            string directoryPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), Consts.Constants.ToolsName);

            if (!Directory.Exists(directoryPath))
                Directory.CreateDirectory(directoryPath);

            _filePath = Path.Combine(directoryPath, Consts.Constants.ConfigFileName);

            if (!File.Exists(_filePath))            
                LoadDefaultSettings();                
            else
                LoadSettingsFromFile();        
        }

        private void LoadSettingsFromFile()
        {
            XDocument xdoc = XDocument.Load(_filePath);

            try
            {
                this.NotebookId_Bible = xdoc.Root.XPathSelectElement(Consts.Constants.ParameterName_NotebookIdBible).Value;
                this.NotebookId_BibleComments = xdoc.Root.XPathSelectElement(Consts.Constants.ParameterName_NotebookIdBibleComments).Value;
                this.NotebookId_BibleStudy = xdoc.Root.XPathSelectElement(Consts.Constants.ParameterName_NotebookIdBibleStudy).Value;
                this.SectionGroupId_Bible = xdoc.Root.XPathSelectElement(Consts.Constants.ParameterName_SectionGroupIdBible).Value;
                this.SectionGroupId_BibleComments = xdoc.Root.XPathSelectElement(Consts.Constants.ParameterName_SectionGroupIdBibleComments).Value;
                this.SectionGroupId_BibleStudy = xdoc.Root.XPathSelectElement(Consts.Constants.ParameterName_SectionGroupIdBibleStudy).Value;
                this.PageName_DefaultBookOverview = xdoc.Root.XPathSelectElement(Consts.Constants.ParameterName_PageNameDefaultBookOverview).Value;
                this.PageName_DefaultComments = xdoc.Root.XPathSelectElement(Consts.Constants.ParameterName_PageNameDefaultComments).Value;
                this.PageName_Notes = xdoc.Root.XPathSelectElement(Consts.Constants.ParameterName_PageNamePageName_Notes).Value;

                XElement lastNotesLinkTimeElement = xdoc.Root.XPathSelectElement(Consts.Constants.ParameterName_LastNotesLinkTime);
                this.LastNotesLinkTime = (lastNotesLinkTimeElement != null && !string.IsNullOrEmpty(lastNotesLinkTimeElement.Value)) 
                                            ? (DateTime?)DateTime.Parse(lastNotesLinkTimeElement.Value) : null;

                XElement newVersionOnServerElement = xdoc.Root.XPathSelectElement(Consts.Constants.ParameterName_NewVersionOnServer);
                this.NewVersionOnServer = (newVersionOnServerElement != null && !string.IsNullOrEmpty(newVersionOnServerElement.Value))
                                            ? new Version(newVersionOnServerElement.Value) : null;

                XElement newVersionOnServerLatestCheckTimeElement = xdoc.Root.XPathSelectElement(Consts.Constants.ParameterName_NewVersionOnServerLatestCheckTime);
                this.NewVersionOnServerLatestCheckTime = (newVersionOnServerLatestCheckTimeElement != null && !string.IsNullOrEmpty(newVersionOnServerLatestCheckTimeElement.Value))
                                            ? (DateTime?)DateTime.Parse(newVersionOnServerLatestCheckTimeElement.Value) : null;
            }
            catch (Exception ex)
            {
                Logger.LogError(ex.Message);
                LoadDefaultSettings();
            }
        }

        private void LoadDefaultSettings()
        {                       
            this.PageName_DefaultBookOverview = Consts.Constants.DefaultPageNameDefaultBookOverview;
            this.PageName_DefaultComments = Consts.Constants.DefaultPageNameDefaultComments;
            this.PageName_Notes = Consts.Constants.DefaultPageName_Notes;
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
                                  new XElement(Consts.Constants.ParameterName_NotebookIdBibleStudy, this.NotebookId_BibleStudy),                                  
                                  new XElement(Consts.Constants.ParameterName_SectionGroupIdBible, this.SectionGroupId_Bible),
                                  new XElement(Consts.Constants.ParameterName_SectionGroupIdBibleComments, this.SectionGroupId_BibleComments),
                                  new XElement(Consts.Constants.ParameterName_SectionGroupIdBibleStudy, this.SectionGroupId_BibleStudy),
                                  new XElement(Consts.Constants.ParameterName_PageNameDefaultBookOverview, this.PageName_DefaultBookOverview),
                                  new XElement(Consts.Constants.ParameterName_PageNameDefaultComments, this.PageName_DefaultComments),
                                  new XElement(Consts.Constants.ParameterName_PageNamePageName_Notes, this.PageName_Notes),
                                  new XElement(Consts.Constants.ParameterName_LastNotesLinkTime, this.LastNotesLinkTime.HasValue 
                                                ? this.LastNotesLinkTime.Value.ToString() : string.Empty),
                                  new XElement(Consts.Constants.ParameterName_NewVersionOnServer, this.NewVersionOnServer),
                                  new XElement(Consts.Constants.ParameterName_NewVersionOnServerLatestCheckTime, this.NewVersionOnServerLatestCheckTime.HasValue
                                                ? this.NewVersionOnServerLatestCheckTime.Value.ToString() : string.Empty)
                                  );

                    xDoc.Save(sw);
                    sw.Flush();                    
                }
            }
        }                 
    }
}


