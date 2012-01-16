﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using BibleCommon.Services;
using BibleCommon.Helpers;
using System.Xml;
using Microsoft.Office.Interop.OneNote;
using System.Xml.Linq;
using System.Xml.XPath;
using System.IO;
using System.Diagnostics;
using BibleCommon;

namespace BibleConfigurator
{
    public partial class MainForm : Form
    {
        private const string SingleNotebookDefaultName = "Holy Bible";
        private const string BibleNotebookDefaultName = "Библия";
        private const string BibleCommentsNotebookDefaultName = "Комментарии к Библии";
        private const string BibleStudyNotebookDefaultName = "Изучение Библии";

        private const string SingleNotebookTemplateFileName = "Holy Bible.onepkg";
        private const string BibleNotebookTemplateFileName = "Библия.onepkg";
        private const string BibleCommentsNotebookTemplateFileName = "Комментарии к Библии.onepkg";
        private const string BibleStudyNotebookTemplateFileName = "Изучение Библии.onepkg";
        private const string TemplatesDirectory = "OneNoteTemplates";

        private Microsoft.Office.Interop.OneNote.Application _oneNoteApp = new Microsoft.Office.Interop.OneNote.Application();

        private string SingleNotebookFromTemplatePath { get; set; }
        private string BibleNotebookFromTemplatePath { get; set; }
        private string BibleCommentsNotebookFromTemplatePath { get; set; }
        private string BibleStudyNotebookFromTemplatePath { get; set; }

        private enum NotebookType
        {
            Single,
            Bible,
            BibleComments,
            BibleStudy
        }

        public MainForm()
        {
            InitializeComponent();                     
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            Logger.Initialize();       

            if (rbSingleNotebook.Checked)
            {
                Settings.Default.NotebookName_Bible = string.Empty;
                Settings.Default.NotebookName_BibleComments = string.Empty;
                Settings.Default.NotebookName_BibleStudy = string.Empty;

                if (chkCreateSingleNotebookFromTemplate.Checked)
                {
                    string notebookName = CreateNotebookFromTemplate(SingleNotebookTemplateFileName, SingleNotebookFromTemplatePath);
                    if (!string.IsNullOrEmpty(notebookName))
                        Settings.Default.NotebookName_Single = notebookName;
                }
                else
                {
                    if (CheckNotebook((string)cbSingleNotebook.SelectedItem, NotebookType.Single))
                        Settings.Default.NotebookName_Single = (string)cbSingleNotebook.SelectedItem;
                }
            }
            else
            {
                Settings.Default.NotebookName_Single = string.Empty;
                Settings.Default.SectionGroupName_Bible = string.Empty;
                Settings.Default.SectionGroupName_BibleComments = string.Empty;
                Settings.Default.SectionGroupName_BibleStudy = string.Empty;

                if (chkCreateBibleNotebookFromTemplate.Checked)
                {
                    string notebookName = CreateNotebookFromTemplate(BibleNotebookTemplateFileName, BibleNotebookFromTemplatePath);
                    if (!string.IsNullOrEmpty(notebookName))
                        Settings.Default.NotebookName_Bible = notebookName;
                }
                else
                {
                    if (CheckNotebook((string)cbBibleNotebook.SelectedItem, NotebookType.Bible))
                        Settings.Default.NotebookName_Bible = (string)cbBibleNotebook.SelectedItem;
                }

                if (chkCreateBibleCommentsNotebookFromTemplate.Checked)
                {
                    string notebookName = CreateNotebookFromTemplate(BibleCommentsNotebookTemplateFileName, BibleCommentsNotebookFromTemplatePath);
                    if (!string.IsNullOrEmpty(notebookName))
                        Settings.Default.NotebookName_BibleComments = notebookName;
                }
                else
                {
                    if (CheckNotebook((string)cbBibleCommentsNotebook.SelectedItem, NotebookType.BibleComments))
                        Settings.Default.NotebookName_BibleComments = (string)cbBibleCommentsNotebook.SelectedItem;
                }

                if (chkCreateBibleStudyNotebookFromTemplate.Checked)
                {
                    string notebookName = CreateNotebookFromTemplate(BibleStudyNotebookTemplateFileName, BibleStudyNotebookFromTemplatePath);
                    if (!string.IsNullOrEmpty(notebookName))
                        Settings.Default.NotebookName_BibleStudy = notebookName;
                }
                else
                {
                    if (CheckNotebook((string)cbBibleStudyNotebook.SelectedItem, NotebookType.BibleStudy))
                        Settings.Default.NotebookName_BibleStudy = (string)cbBibleStudyNotebook.SelectedItem;
                }
            }

            if (!Logger.WasErrorLogged)
            {
                Close();
                Settings.Default.Save();
            }   
        }

        private bool CheckNotebook(string notebookName, NotebookType notebookType)
        {
            string notebookId = OneNoteUtils.GetNotebookId(_oneNoteApp, notebookName);
            string errorText = string.Empty;

            if (!string.IsNullOrEmpty(notebookId))
            {
                string xml;
                XmlNamespaceManager xnm;
                _oneNoteApp.GetHierarchy(notebookId, HierarchyScope.hsSections, out xml);
                XDocument notebookDoc = OneNoteUtils.GetXDocument(xml, out xnm);

                switch (notebookType)
                {
                    case NotebookType.Single:
                        XElement bibleSectionGroup = notebookDoc.Root.XPathSelectElement(string.Format("one:SectionGroup[@name='{0}']", Settings.Default.SectionGroupName_Bible), xnm);
                        if (bibleSectionGroup != null)
                        {
                            XElement bibleCommentsSectionGroup = notebookDoc.Root.XPathSelectElement(string.Format("one:SectionGroup[@name='{0}']", Settings.Default.SectionGroupName_BibleComments), xnm);
                            if (bibleCommentsSectionGroup != null)
                            {
                                XElement bibleStudySectionGroup = notebookDoc.Root.XPathSelectElement(string.Format("one:SectionGroup[@name='{0}']", Settings.Default.SectionGroupName_BibleStudy), xnm);
                                if (bibleStudySectionGroup != null)
                                {
                                    return true;
                                }
                            }
                        }
                        //else
                        //    errorText = string.Format("", )
                        break;
                }                
            }
            else
                Logger.LogError(string.Format("Не найдена записная книжка '{0}'.", notebookName));

            Logger.LogError(string.Format("Указана неподходящая записная книжка '{0}': {1}.", notebookName, errorText));
            return false;
        }

        private string CreateNotebookFromTemplate(string notebookTemplateFileName, string notebookFromTemplatePath)
        {
            string s;
            string packageFilePath = Path.Combine(Path.Combine(Path.GetDirectoryName(Utils.GetCurrentDirectory()), TemplatesDirectory), notebookTemplateFileName);

            if (File.Exists(packageFilePath))
            {
                string folderPath = Path.Combine(notebookFromTemplatePath, Path.GetFileNameWithoutExtension(notebookTemplateFileName));                

                folderPath = GetNewDirectoryPath(folderPath);

                if (!string.IsNullOrEmpty(folderPath))
                {
                    _oneNoteApp.OpenPackage(packageFilePath, folderPath, out s);

                    string[] files = Directory.GetFiles(s, "*.onetoc2", SearchOption.TopDirectoryOnly);
                    if (files.Length > 0)
                        Process.Start(files[0]);
                    else
                        Logger.LogError(string.Format("Ошибка при открытии записной книжки '{0}'.", notebookTemplateFileName));

                    return Path.GetFileNameWithoutExtension(folderPath);
                }
                else
                    Logger.LogError("Не удаётся создать записную книжку. Выберите другую папку.");
            }
            else
                Logger.LogError(string.Format("Не найден шаблон записной книжки по адресу '{0}'.", packageFilePath));

            return string.Empty;
        }

        private string GetNewDirectoryPath(string folderPath)
        {
            string result = folderPath;
            for (int i = 0; i < 100; i++)
            {
                result = folderPath + (i > 0 ? " (" + i.ToString() + ")" : string.Empty);

                if (!Directory.Exists(result))
                    return result;
            }

            return string.Empty;
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            PrepareFolderBrowser();            

            rbSingleNotebook.Checked = (Settings.Default.NotebookName_Bible == Settings.Default.NotebookName_BibleComments)
                                    && (Settings.Default.NotebookName_Bible == Settings.Default.NotebookName_BibleStudy);

            rbMultiNotebook.Checked = !rbSingleNotebook.Checked;
            rbMultiNotebook_CheckedChanged(this, null);


            Dictionary<string, string> notebooks = GetNotebooks();

            cbSingleNotebook.DataSource = notebooks.Keys.ToList();
            cbBibleNotebook.DataSource = notebooks.Keys.ToList();
            cbBibleCommentsNotebook.DataSource = notebooks.Keys.ToList();
            cbBibleStudyNotebook.DataSource = notebooks.Keys.ToList();

            cbSingleNotebook.SelectedItem = !string.IsNullOrEmpty(Settings.Default.NotebookName_Single) ?
                Settings.Default.NotebookName_Single : SingleNotebookDefaultName;            
            cbBibleNotebook.SelectedItem = !string.IsNullOrEmpty(Settings.Default.NotebookName_Bible) ?
                Settings.Default.NotebookName_Bible : BibleNotebookDefaultName;            
            cbBibleCommentsNotebook.SelectedItem = !string.IsNullOrEmpty(Settings.Default.NotebookName_BibleComments) ?
                Settings.Default.NotebookName_BibleComments : BibleCommentsNotebookDefaultName;            
            cbBibleStudyNotebook.SelectedItem = !string.IsNullOrEmpty(Settings.Default.NotebookName_BibleStudy) ?
                Settings.Default.NotebookName_BibleStudy : BibleStudyNotebookDefaultName;
        }

        private void PrepareFolderBrowser()
        {
            string[] directories = Directory.GetDirectories(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                                        "*OneNote*", SearchOption.TopDirectoryOnly);
            if (directories.Length > 0)
                folderBrowserDialog.SelectedPath = directories[0];

            //folderBrowserDialog.RootFolder = Environment.SpecialFolder.MyDocuments;

            folderBrowserDialog.Description = "Укажите расположение записной книжки";
            folderBrowserDialog.ShowNewFolderButton = true;            
        }

        public Dictionary<string, string> GetNotebooks()
        {
            Dictionary<string, string> result = new Dictionary<string, string>();

            string xml;
            XmlNamespaceManager xnm;
            _oneNoteApp.GetHierarchy(null, HierarchyScope.hsNotebooks, out xml);
            XDocument doc = OneNoteUtils.GetXDocument(xml, out xnm);
            foreach (XElement notebook in doc.Root.XPathSelectElements("one:Notebook", xnm))
            {
                string name = (string)notebook.Attribute("name");
                string id = (string)notebook.Attribute("ID");
                result.Add(name, id);
            }

            return result;
        }

        private void rbMultiNotebook_CheckedChanged(object sender, EventArgs e)
        {
            cbSingleNotebook.Enabled = rbSingleNotebook.Checked;
            btnSingleNotebookParameters.Enabled = rbSingleNotebook.Checked;
            chkCreateSingleNotebookFromTemplate.Enabled = rbSingleNotebook.Checked;
            btnSingleNotebookParameters.Enabled = rbSingleNotebook.Checked;

            cbBibleNotebook.Enabled = rbMultiNotebook.Checked;
            cbBibleCommentsNotebook.Enabled = rbMultiNotebook.Checked;
            cbBibleStudyNotebook.Enabled = rbMultiNotebook.Checked;
            chkCreateBibleNotebookFromTemplate.Enabled = rbMultiNotebook.Checked;
            chkCreateBibleCommentsNotebookFromTemplate.Enabled = rbMultiNotebook.Checked;
            chkCreateBibleStudyNotebookFromTemplate.Enabled = rbMultiNotebook.Checked;
        }

        private void chkCreateSingleNotebookFromTemplate_CheckedChanged(object sender, EventArgs e)
        {
            if (chkCreateSingleNotebookFromTemplate.Checked)
            {
                if (folderBrowserDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    SingleNotebookFromTemplatePath = folderBrowserDialog.SelectedPath;
                }
                else
                    chkCreateSingleNotebookFromTemplate.Checked = false;
            }

            cbSingleNotebook.Enabled = !chkCreateSingleNotebookFromTemplate.Checked;
            btnSingleNotebookParameters.Enabled = !chkCreateSingleNotebookFromTemplate.Checked;
        }

        private void chkCreateBibleNotebookFromTemplate_CheckedChanged(object sender, EventArgs e)
        {
            if (chkCreateBibleNotebookFromTemplate.Checked)
            {
                if (folderBrowserDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    BibleNotebookFromTemplatePath = folderBrowserDialog.SelectedPath;
                }
                else
                    chkCreateBibleNotebookFromTemplate.Checked = false;
            }

            cbBibleNotebook.Enabled = !chkCreateBibleNotebookFromTemplate.Checked;
        }

        private void chkCreateBibleCommentsNotebookFromTemplate_CheckedChanged(object sender, EventArgs e)
        {
            if (chkCreateBibleCommentsNotebookFromTemplate.Checked)
            {
                if (folderBrowserDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    BibleCommentsNotebookFromTemplatePath = folderBrowserDialog.SelectedPath;
                }
                else
                    chkCreateBibleCommentsNotebookFromTemplate.Checked = false;
            }

            cbBibleCommentsNotebook.Enabled = !chkCreateBibleCommentsNotebookFromTemplate.Checked;
        }

        private void chkCreateBibleStudyNotebookFromTemplate_CheckedChanged(object sender, EventArgs e)
        {
            if (chkCreateBibleStudyNotebookFromTemplate.Checked)
            {
                if (folderBrowserDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    BibleStudyNotebookFromTemplatePath = folderBrowserDialog.SelectedPath;
                }
                else
                    chkCreateBibleStudyNotebookFromTemplate.Checked = false;
            }

            cbBibleStudyNotebook.Enabled = !chkCreateBibleStudyNotebookFromTemplate.Checked;
        }

        private void btnSingleNotebookParameters_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty((string)cbSingleNotebook.SelectedItem))
            {
                NotebookParametersForm notebookParametersForm = new NotebookParametersForm(_oneNoteApp, (string)cbSingleNotebook.SelectedItem);
                if (notebookParametersForm.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    //либо переименовать, либо запомнить дефолтные
                }
            }
            else
            {
                Logger.LogMessage("Не указана записная книжка.");
            }
        }
    }
}
