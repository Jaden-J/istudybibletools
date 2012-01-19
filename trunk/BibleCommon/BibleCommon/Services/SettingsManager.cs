using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Xml.Linq;
using System.Xml.XPath;
using BibleCommon.Helpers;
using Microsoft.Office.Interop.OneNote;

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

        public string NotebookId_Single { get; private set; }
        public string NotebookId_Bible { get; private set; }
        public string NotebookId_BibleComments { get; private set; }
        public string NotebookId_BibleStudy { get; private set; }
        public string NotebookName_Single { get; private set; }
        public string NotebookName_Bible { get; private set; }
        public string NotebookName_BibleComments { get; private set; }
        public string NotebookName_BibleStudy { get; private set; }
        public string SectionGroupName_Bible { get; private set; }
        public string SectionGroupName_BibleComments { get; private set; }
        public string SectionGroupName_BibleStudy { get; private set; }
        public string PageName_DefaultDescription { get; private set; }
        public string PageName_DefaultBookOverview { get; private set; }
        public string PageName_Notes { get; private set; }

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

            LoadNotebookIds();
        }

        private void LoadSettingsFromFile()
        {
            XDocument xdoc = XDocument.Load(_filePath);

            this.NotebookName_Bible = xdoc.Root.XPathSelectElement(Consts.Constants.ParameterName_NotebookNameBible).Value;
            this.NotebookName_BibleComments = xdoc.Root.XPathSelectElement(Consts.Constants.ParameterName_NotebookNameBibleComments).Value;
            this.NotebookName_BibleStudy = xdoc.Root.XPathSelectElement(Consts.Constants.ParameterName_NotebookNameBible).Value;
            this.NotebookName_Single = xdoc.Root.XPathSelectElement(Consts.Constants.ParameterName_NotebookNameSingle).Value;
            this.SectionGroupName_Bible = xdoc.Root.XPathSelectElement(Consts.Constants.ParameterName_SectionGroupNameBible).Value;
            this.SectionGroupName_BibleComments = xdoc.Root.XPathSelectElement(Consts.Constants.ParameterName_SectionGroupNameBibleComments).Value;
            this.SectionGroupName_BibleStudy = xdoc.Root.XPathSelectElement(Consts.Constants.ParameterName_SectionGroupNameBibleStudy).Value;
            this.PageName_DefaultBookOverview = xdoc.Root.XPathSelectElement(Consts.Constants.ParameterName_PageNameDefaultBookOverview).Value;
            this.PageName_DefaultDescription = xdoc.Root.XPathSelectElement(Consts.Constants.ParameterName_PageNameDefaultDescription).Value;
            this.PageName_Notes = xdoc.Root.XPathSelectElement(Consts.Constants.ParameterName_PageNamePageName_Notes).Value;
        }

        private void LoadNotebookIds()
        {
            Application oneNoteApp = new Application();
            this.NotebookId_Bible = GetNotebookId(oneNoteApp, this.NotebookName_Bible);
            this.NotebookId_BibleComments = GetNotebookId(oneNoteApp, this.NotebookName_BibleComments);
            this.NotebookId_BibleStudy = GetNotebookId(oneNoteApp, this.NotebookName_BibleStudy);
        }

        private void LoadDefaultSettings()
        {
            this.NotebookName_Single = string.Empty;
            this.NotebookName_Bible = Consts.Constants.DefaultNotebookNameBible;
            this.NotebookName_BibleComments = Consts.Constants.DefaultNotebookNameBibleComments;
            this.NotebookName_BibleStudy = Consts.Constants.DefaultNotebookNameBibleStudy;

            this.SectionGroupName_Bible = string.Empty;
            this.SectionGroupName_BibleComments = string.Empty;
            this.SectionGroupName_BibleStudy = string.Empty;

            this.PageName_DefaultBookOverview = Consts.Constants.DefaultPageNameDefaultBookOverview;
            this.PageName_DefaultDescription = Consts.Constants.DefaultPageNameDefaultDescription;
            this.PageName_Notes = Consts.Constants.DefaultPageName_Notes;
        }

        public void SaveSettings()
        {
            // сохранять идентификаторы, и просто преолбразовывать при чтении их в имена (и зхаписные книжки, и группы разделов
            using (FileStream fs = new FileStream(
            XDocument xDoc = XDocument.Parse("<Settings></Settings>");
            xDoc.Root.Add(new XElement(Consts.Constants.ParameterName_NotebookNameBible, this.NotebookName_Bible),
                          new XElement(Consts.Constants.ParameterName_NotebookNameBibleComments, this.NotebookName_BibleComments),
                          new XElement(Consts.Constants.ParameterName_NotebookNameBibleStudy, this.NotebookName_BibleStudy),
                          new XElement(Consts.Constants.ParameterName_NotebookNameSingle, this.NotebookName_Single),
                          new XElement(Consts.Constants.ParameterName_SectionGroupNameBible, this.SectionGroupName_Bible),
                          new XElement(Consts.Constants.ParameterName_SectionGroupNameBibleComments, this.SectionGroupName_BibleComments),
                          new XElement(Consts.Constants.ParameterName_SectionGroupNameBibleStudy, this.SectionGroupName_BibleStudy),
                          new XElement(Consts.Constants.ParameterName_PageNameDefaultBookOverview, this.PageName_DefaultBookOverview),
                          new XElement(Consts.Constants.ParameterName_PageNameDefaultDescription, this.PageName_DefaultDescription),
                          new XElement(Consts.Constants.ParameterName_PageNamePageName_Notes, this.PageName_Notes)
                          );
            

        }

        private string GetNotebookId(Application oneNoteApp, string notebookName)
        {
            string result = string.Empty;

            if (!string.IsNullOrEmpty(notebookName))
            {
                 result = OneNoteUtils.GetNotebookId(oneNoteApp, notebookName);

                if (string.IsNullOrEmpty(result))
                    throw new Exception(string.Format("Не найдено записной книжки: {0}", notebookName));

            }

            return result;
        }       
    }
}


