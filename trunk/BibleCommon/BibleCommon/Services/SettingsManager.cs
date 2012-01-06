using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Xml.Linq;
using System.Xml.XPath;
using BibleCommon.Helpers;

namespace BibleCommon.Services
{
    public class SettingsManager
    {
        internal const string SettingsFileName = "settings.config";
        private static object _locker = new object();

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

        public string NotebookName { get; private set; }
        public string BibleSectionGroupName { get; private set; }
        public string StudyBibleSectionGroupName { get; private set; }
        public string ResearchSectionGroupName { get; private set; }
        public string DescriptionPageDefaultName { get; private set; }
        public string BookOverviewPageDefaultName { get; private set; }
        public string NotesPageDefaultName { get; private set; }    

        protected SettingsManager()
        {            
            string currentDirectory = Utils.GetCurrentDirectory();
            string filePath = Path.Combine(currentDirectory, SettingsFileName);
            if (!ReadSettingsFile(filePath))
            {
                filePath = Path.Combine(Directory.GetParent(currentDirectory).ToString(), SettingsFileName);
                if (!ReadSettingsFile(filePath))
                {
                    throw new Exception(string.Format("Не найден файл настроек '{0}'.", SettingsFileName));
                }
            }
        }

        private bool ReadSettingsFile(string filePath)
        {
            if (File.Exists(filePath))
            {
                try
                {
                    XDocument xdoc = XDocument.Load(filePath);

                    this.NotebookName = xdoc.Root.XPathSelectElement("NotebookName").Value;
                    this.BibleSectionGroupName = xdoc.Root.XPathSelectElement("BibleSectionGroupName").Value;
                    this.StudyBibleSectionGroupName = xdoc.Root.XPathSelectElement("StudyBibleSectionGroupName").Value;
                    this.ResearchSectionGroupName = xdoc.Root.XPathSelectElement("ResearchSectionGroupName").Value;
                    this.DescriptionPageDefaultName = xdoc.Root.XPathSelectElement("DescriptionPageDefaultName").Value;
                    this.BookOverviewPageDefaultName = xdoc.Root.XPathSelectElement("BookOverviewPageDefaultName").Value;
                    this.NotesPageDefaultName = xdoc.Root.XPathSelectElement("NotesPageDefaultName").Value;
                }
                catch (Exception ex)
                {
                    throw new Exception(string.Format("Ошибка при чтении файла настроек: {0}", ex.Message));
                }

                return true;
            }

            return false;
        }
    }
}
