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

        public string NotebookId_Bible { get; private set; }
        public string NotebookId_BibleComments { get; private set; }
        public string NotebookId_BibleStudy { get; private set; }
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
            this.NotebookName_Bible = Settings.Default.NotebookName_Bible;
            this.NotebookName_BibleComments = Settings.Default.NotebookName_BibleComments;
            this.NotebookName_BibleStudy = Settings.Default.NotebookName_BibleStudy;

            this.SectionGroupName_Bible = Settings.Default.SectionGroupName_Bible;
            this.SectionGroupName_BibleComments = Settings.Default.SectionGroupName_BibleComments;
            this.SectionGroupName_BibleStudy = Settings.Default.SectionGroupName_BibleStudy;

            this.PageName_DefaultBookOverview = Settings.Default.PageName_DefaultBookOverview;
            this.PageName_DefaultDescription = Settings.Default.PageName_DefaultDescription;
            this.PageName_Notes = Settings.Default.PageName_Notes;

            Application oneNoteApp = new Application();
            this.NotebookId_Bible = GetNotebookId(oneNoteApp, this.NotebookName_Bible);
            this.NotebookId_BibleComments = GetNotebookId(oneNoteApp, this.NotebookName_BibleComments);
            this.NotebookId_BibleStudy = GetNotebookId(oneNoteApp, this.NotebookName_BibleStudy);
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


