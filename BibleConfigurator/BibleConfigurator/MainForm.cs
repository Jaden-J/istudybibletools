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
using System.Threading;
using BibleConfigurator.Tools;

namespace BibleConfigurator
{
    public partial class MainForm : Form
    {
        private Microsoft.Office.Interop.OneNote.Application _oneNoteApp = new Microsoft.Office.Interop.OneNote.Application();

        private string SingleNotebookFromTemplatePath { get; set; }
        private string BibleNotebookFromTemplatePath { get; set; }
        private string BibleCommentsNotebookFromTemplatePath { get; set; }
        private string BibleStudyNotebookFromTemplatePath { get; set; }

        private bool _wasSearchedSectionGroupsInSingleNotebook = false;       
        

        private const int LoadParametersAttemptsCount = 40;         // количество попыток загрузки параметров после команды создания записных книжек из шаблона
        private const int LoadParametersPauseBetweenAttempts = 10;             // количество секунд ожидания между попытками загрузки параметров
        private const string LoadParametersImageFileName = "loader.gif";


        private NotebookParametersForm _notebookParametersForm = null;

        public MainForm()
        {
            InitializeComponent();
            BibleCommon.Services.Logger.Init("BibleConfigurator");
        }

        public bool StopExternalProcess { get; set; }
        

        private void btnOK_Click(object sender, EventArgs e)
        {          
            btnOK.Enabled = false;

            try
            {
                Logger.Initialize();

                string notebookId;
                string notebookName;

                if (rbSingleNotebook.Checked)
                {
                    if (chkCreateSingleNotebookFromTemplate.Checked)
                    {
                        notebookName = CreateNotebookFromTemplate(Consts.SingleNotebookTemplateFileName, SingleNotebookFromTemplatePath);
                        if (!string.IsNullOrEmpty(notebookName))
                        {
                            WaitAndLoadParameters(NotebookType.Single, notebookName);
                            SearchForCorrespondenceSectionGroups(SettingsManager.Instance.NotebookId_Bible);
                        }
                    }
                    else
                    {
                        notebookName = (string)cbSingleNotebook.SelectedItem;                        
                        if (TryToLoadNotebookParameters(NotebookType.Single, notebookName, out notebookId))
                        {
                            if (_notebookParametersForm != null && _notebookParametersForm.RenamedSectionGroups.Count > 0)
                                RenameSectionGroupsForm(notebookId, _notebookParametersForm.RenamedSectionGroups);

                            if (!_wasSearchedSectionGroupsInSingleNotebook)
                            {
                                try
                                {
                                    SearchForCorrespondenceSectionGroups(notebookId);
                                }
                                catch (InvalidNotebookException)
                                {
                                    Logger.LogError("Указана неподходящая записная книжка.");
                                }
                            }
                        }
                        
                    }
                }
                else
                {
                    SettingsManager.Instance.SectionGroupId_Bible = string.Empty;
                    SettingsManager.Instance.SectionGroupId_BibleComments = string.Empty;
                    SettingsManager.Instance.SectionGroupId_BibleStudy = string.Empty;

                    if (chkCreateBibleStudyNotebookFromTemplate.Checked)
                    {
                        notebookName = CreateNotebookFromTemplate(Consts.BibleStudyNotebookTemplateFileName, BibleStudyNotebookFromTemplatePath);
                        if (!string.IsNullOrEmpty(notebookName))                        
                            WaitAndLoadParameters(NotebookType.BibleStudy, notebookName);                         // выйдем из метода только когда OneNote отработает                        
                    }
                    else
                    {
                        notebookName = (string)cbBibleStudyNotebook.SelectedItem;
                        TryToLoadNotebookParameters(NotebookType.BibleStudy, notebookName, out notebookId);
                    }

                    if (!Logger.WasErrorLogged)
                    {
                        if (chkCreateBibleCommentsNotebookFromTemplate.Checked)
                        {
                            notebookName = CreateNotebookFromTemplate(Consts.BibleCommentsNotebookTemplateFileName, BibleCommentsNotebookFromTemplatePath);
                            if (!string.IsNullOrEmpty(notebookName))
                                WaitAndLoadParameters(NotebookType.BibleComments, notebookName);                         // выйдем из метода только когда OneNote отработает                                                
                        }
                        else
                        {
                            notebookName = (string)cbBibleCommentsNotebook.SelectedItem;
                            TryToLoadNotebookParameters(NotebookType.BibleComments, notebookName, out notebookId);
                        }

                        if (!Logger.WasErrorLogged)
                        {
                            if (chkCreateBibleNotebookFromTemplate.Checked)  // записную книжку для Библии создаём в самом конце, так как она дольше всех создаётся
                            {
                                notebookName = CreateNotebookFromTemplate(Consts.BibleNotebookTemplateFileName, BibleNotebookFromTemplatePath);
                                if (!string.IsNullOrEmpty(notebookName))
                                    WaitAndLoadParameters(NotebookType.Bible, notebookName);                         // выйдем из метода только когда OneNote отработает
                            }
                            else
                            {
                                notebookName = (string)cbBibleNotebook.SelectedItem;
                                TryToLoadNotebookParameters(NotebookType.Bible, notebookName, out notebookId);
                            }
                        }
                    }
                }

                if (!Logger.WasErrorLogged)
                {
                    SetPageNameParameters();

                    SettingsManager.Instance.Save();
                    Close();
                }

            }
            catch (LoadParametersException)
            {
                Logger.LogError("Не удалось запросить данные о записных книжках из OneNote. Повторите операцию.");
                LoadParameters();                
            }
            finally
            {
                btnOK.Enabled = true;
            }
        }

        private void SetPageNameParameters()
        {
            if (chkDefaultPageNameParameters.Checked)
            {
                SettingsManager.Instance.PageName_DefaultBookOverview = Consts.PageNameDefaultBookOverview;
                SettingsManager.Instance.PageName_Notes = Consts.PageNameNotes;
                SettingsManager.Instance.PageName_DefaultComments = Consts.PageNameDefaultComments;
            }
            else
            {
                if (!string.IsNullOrEmpty(tbBookOverviewName.Text))
                    SettingsManager.Instance.PageName_DefaultBookOverview = tbBookOverviewName.Text;

                if (!string.IsNullOrEmpty(tbNotesPageName.Text))
                    SettingsManager.Instance.PageName_Notes = tbNotesPageName.Text;

                if (!string.IsNullOrEmpty(tbPageDescriptionName.Text))
                    SettingsManager.Instance.PageName_DefaultComments = tbPageDescriptionName.Text;
            }
        }

        private void WaitAndLoadParameters(NotebookType notebookType, string notebookName)
        {   
            PrepareForExternalProcessing(100, 1, string.Format("Создание записной книжки '{0}'", notebookName));
            
            bool parametersWasLoad = false;

            try
            {
                string notebookId;                
                for (int i = 0; i <= LoadParametersAttemptsCount; i++)
                {
                    pbMain.PerformStep();
                    System.Windows.Forms.Application.DoEvents();

                    if (TryToLoadNotebookParameters(notebookType, notebookName, out notebookId, true))
                    {
                        parametersWasLoad = true;
                        break;
                    }

                    Thread.Sleep(LoadParametersPauseBetweenAttempts * 1000);
                }                
            }
            finally
            {
                ExternalProcessingDone(string.Empty);                
            }

            if (!parametersWasLoad)
                throw new LoadParametersException();
        }

        private bool TryToLoadNotebookParameters(NotebookType notebookType, string notebookName, out string notebookId, bool silientMode = false)
        {
            notebookId = string.Empty;

            try
            {
                notebookId = OneNoteUtils.GetNotebookIdByName(_oneNoteApp, notebookName);                
                if (NotebookChecker.CheckNotebook(_oneNoteApp, notebookId, notebookType))
                {
                    switch (notebookType)
                    {
                        case NotebookType.Single:
                            SettingsManager.Instance.NotebookId_Bible = notebookId;
                            SettingsManager.Instance.NotebookId_BibleComments = notebookId;
                            SettingsManager.Instance.NotebookId_BibleStudy = notebookId;
                            break;
                        case NotebookType.Bible:
                            SettingsManager.Instance.NotebookId_Bible = notebookId;
                            break;
                        case NotebookType.BibleComments:
                            SettingsManager.Instance.NotebookId_BibleComments = notebookId;
                            break;
                        case NotebookType.BibleStudy:
                            SettingsManager.Instance.NotebookId_BibleStudy = notebookId;
                            break;
                    }

                    return true;
                }
                else
                {
                    if (!silientMode)
                        Logger.LogError(string.Format("Указана неподходящая записная книжка '{0}'.", notebookName));
                }
            }
            catch (Exception ex)
            {
                BibleCommon.Services.Logger.LogError(ex);
            }

            return false;
        }

        private void SearchForCorrespondenceSectionGroups(string notebookId)
        {
            string xml;
            XmlNamespaceManager xnm;
            _oneNoteApp.GetHierarchy(notebookId, HierarchyScope.hsSections, out xml);
            XDocument notebookDoc = OneNoteUtils.GetXDocument(xml, out xnm);

            List<SectionGroupType> sectionGroups = new List<SectionGroupType>();

            foreach (XElement sectionGroup in notebookDoc.Root.XPathSelectElements("one:SectionGroup", xnm))
            {
                string id = (string)sectionGroup.Attribute("ID");

                if (NotebookChecker.ElementIsBible(sectionGroup, xnm) && !sectionGroups.Contains(SectionGroupType.Bible))
                {
                    SettingsManager.Instance.SectionGroupId_Bible = id;
                    sectionGroups.Add(SectionGroupType.Bible);
                }
                else if (NotebookChecker.ElementIsBibleComments(sectionGroup, xnm) && !sectionGroups.Contains(SectionGroupType.BibleComments))
                {
                    SettingsManager.Instance.SectionGroupId_BibleComments = id;
                    sectionGroups.Add(SectionGroupType.BibleComments);
                }
                else if (!sectionGroups.Contains(SectionGroupType.BibleStudy))
                {
                    SettingsManager.Instance.SectionGroupId_BibleStudy = id;
                    sectionGroups.Add(SectionGroupType.BibleStudy);
                }
                else
                    throw new InvalidNotebookException();
            }

            if (sectionGroups.Count < 3)
                throw new InvalidNotebookException();
        }

        private void RenameSectionGroupsForm(string notebookId, Dictionary<string, string> renamedSectionGroups)
        {
            string xml;
            XmlNamespaceManager xnm;
            _oneNoteApp.GetHierarchy(notebookId, HierarchyScope.hsSections, out xml);
            XDocument notebookDoc = OneNoteUtils.GetXDocument(xml, out xnm);

            foreach (string sectionGroupId in renamedSectionGroups.Keys)
            {
                XElement sectionGroup = notebookDoc.Root.XPathSelectElement(string.Format("one:SectionGroup[@ID='{0}']", sectionGroupId), xnm);

                if (sectionGroup != null)
                {
                    sectionGroup.SetAttributeValue("name", renamedSectionGroups[sectionGroupId]);
                }
                else
                    Logger.LogError(string.Format("Не найдена группа разделов '{0}'.", sectionGroupId));
            }

            _oneNoteApp.UpdateHierarchy(notebookDoc.ToString());
        }

        private string CreateNotebookFromTemplate(string notebookTemplateFileName, string notebookFromTemplatePath)
        {
            string s;
            string packageDirectory = Path.Combine(Path.GetDirectoryName(Path.GetDirectoryName(Utils.GetCurrentDirectory())), Consts.TemplatesDirectory);
            string packageFilePath = Path.Combine(packageDirectory, notebookTemplateFileName);

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
            SetNotebooksDefaultPaths();

            LoadParameters();
        }

        private void LoadParameters()
        {
            Dictionary<string, string> notebooks = GetNotebooks();
            string singleNotebookId = SearchForNotebook(notebooks.Keys, NotebookType.Single);
            string bibleNotebookId = SearchForNotebook(notebooks.Keys, NotebookType.Bible);
            string bibleCommentsNotebookId = SearchForNotebook(notebooks.Keys, NotebookType.BibleComments);
            string bibleStudyNotebookId = SearchForNotebook(notebooks.Keys, NotebookType.BibleStudy);

            rbSingleNotebook.Checked = SettingsManager.Instance.NotebookId_Bible == SettingsManager.Instance.NotebookId_BibleComments
                                    && SettingsManager.Instance.NotebookId_Bible == SettingsManager.Instance.NotebookId_BibleStudy
                                    && !string.IsNullOrEmpty(singleNotebookId);

            rbMultiNotebook.Checked = !rbSingleNotebook.Checked;
            rbMultiNotebook_CheckedChanged(this, null);            

            cbSingleNotebook.DataSource = notebooks.Values.ToList();
            cbBibleNotebook.DataSource = notebooks.Values.ToList();
            cbBibleCommentsNotebook.DataSource = notebooks.Values.ToList();
            cbBibleStudyNotebook.DataSource = notebooks.Values.ToList();

            SetNotebookParameters(rbSingleNotebook.Checked, !string.IsNullOrEmpty(singleNotebookId) ? notebooks[singleNotebookId] :  Consts.SingleNotebookDefaultName, 
                notebooks, SettingsManager.Instance.NotebookId_Bible, cbSingleNotebook, chkCreateSingleNotebookFromTemplate);

            SetNotebookParameters(rbMultiNotebook.Checked, !string.IsNullOrEmpty(bibleNotebookId) ? notebooks[bibleNotebookId] :  Consts.BibleNotebookDefaultName, 
                notebooks, SettingsManager.Instance.NotebookId_Bible, cbBibleNotebook, chkCreateBibleNotebookFromTemplate);

            SetNotebookParameters(rbMultiNotebook.Checked, !string.IsNullOrEmpty(bibleCommentsNotebookId) ? notebooks[bibleCommentsNotebookId] :  Consts.BibleCommentsNotebookDefaultName, 
                notebooks, SettingsManager.Instance.NotebookId_BibleComments, cbBibleCommentsNotebook, chkCreateBibleCommentsNotebookFromTemplate);

            SetNotebookParameters(rbMultiNotebook.Checked, !string.IsNullOrEmpty(bibleStudyNotebookId) ? notebooks[bibleStudyNotebookId] :  Consts.BibleStudyNotebookDefaultName, 
                notebooks, SettingsManager.Instance.NotebookId_BibleStudy, cbBibleStudyNotebook, chkCreateBibleStudyNotebookFromTemplate);

            tbBookOverviewName.Text = SettingsManager.Instance.PageName_DefaultBookOverview;
            tbNotesPageName.Text = SettingsManager.Instance.PageName_Notes;
            tbPageDescriptionName.Text = SettingsManager.Instance.PageName_DefaultComments;
        }

        private string SearchForNotebook(IEnumerable<string> notebooksIds, NotebookType notebookType)
        {
            foreach (string notebookId in notebooksIds)
            {
                if (NotebookChecker.CheckNotebook(_oneNoteApp, notebookId, notebookType))
                {
                    return notebookId;
                }
            }

            return null;
        }

        private static void SetNotebookParameters(bool loadNameFromSettings, string defaultName, Dictionary<string, string> notebooks, string notebookId, ComboBox cb, CheckBox chk)
        {
            string notebookName = (loadNameFromSettings && !string.IsNullOrEmpty(notebookId)) ? TryToGetNotebookName(notebooks, notebookId) : defaultName;
            if (cb.Items.Contains(notebookName))
                cb.SelectedItem = notebookName;
            else
                chk.Checked = true;
        }

        private static string TryToGetNotebookName(Dictionary<string, string> notebooks, string notebookId)
        {
            if (notebooks.ContainsKey(notebookId))
                return notebooks[notebookId];

            return string.Empty;
        }

        private void SetNotebooksDefaultPaths()
        {
            // по дефолту пути такие
            SingleNotebookFromTemplatePath = folderBrowserDialog.SelectedPath;
            BibleNotebookFromTemplatePath = folderBrowserDialog.SelectedPath;
            BibleCommentsNotebookFromTemplatePath = folderBrowserDialog.SelectedPath;
            BibleStudyNotebookFromTemplatePath = folderBrowserDialog.SelectedPath;
        }

        private void PrepareFolderBrowser()
        {
            string myDocumentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string[] directories = Directory.GetDirectories(myDocumentsPath, "*OneNote*", SearchOption.TopDirectoryOnly);
            if (directories.Length > 0)
                folderBrowserDialog.SelectedPath = directories[0];
            else
                folderBrowserDialog.SelectedPath = myDocumentsPath;            

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
                string name = (string)notebook.Attribute("nickname");
                if (string.IsNullOrEmpty(name))
                    name = (string)notebook.Attribute("name");
                string id = (string)notebook.Attribute("ID");
                result.Add(id, name);
            }

            return result;
        }

        private void rbMultiNotebook_CheckedChanged(object sender, EventArgs e)
        {
            cbSingleNotebook.Enabled = rbSingleNotebook.Checked;
            btnSingleNotebookParameters.Enabled = rbSingleNotebook.Checked;
            chkCreateSingleNotebookFromTemplate.Enabled = rbSingleNotebook.Checked;
            btnSingleNotebookParameters.Enabled = rbSingleNotebook.Checked;
            btnSingleNotebookSetPath.Enabled = rbSingleNotebook.Checked;

            cbBibleNotebook.Enabled = rbMultiNotebook.Checked;
            cbBibleCommentsNotebook.Enabled = rbMultiNotebook.Checked;
            cbBibleStudyNotebook.Enabled = rbMultiNotebook.Checked;
            chkCreateBibleNotebookFromTemplate.Enabled = rbMultiNotebook.Checked;
            chkCreateBibleCommentsNotebookFromTemplate.Enabled = rbMultiNotebook.Checked;
            chkCreateBibleStudyNotebookFromTemplate.Enabled = rbMultiNotebook.Checked;
            btnBibleNotebookSetPath.Enabled = rbMultiNotebook.Checked;
            btnBibleCommentsNotebookSetPath.Enabled = rbMultiNotebook.Checked;
            btnBibleStudyNotebookSetPath.Enabled = rbMultiNotebook.Checked;

            if (rbSingleNotebook.Checked)
            {
                chkCreateSingleNotebookFromTemplate_CheckedChanged(this, null);
            }
            else
            {
                chkCreateBibleNotebookFromTemplate_CheckedChanged(this, null);
                chkCreateBibleCommentsNotebookFromTemplate_CheckedChanged(this, null);
                chkCreateBibleStudyNotebookFromTemplate_CheckedChanged(this, null);
            }
        }

        private void chkCreateSingleNotebookFromTemplate_CheckedChanged(object sender, EventArgs e)
        {
            cbSingleNotebook.Enabled = !chkCreateSingleNotebookFromTemplate.Checked;
            btnSingleNotebookParameters.Enabled = !chkCreateSingleNotebookFromTemplate.Checked;
            btnSingleNotebookSetPath.Enabled = chkCreateSingleNotebookFromTemplate.Checked;
        }

        private void chkCreateBibleNotebookFromTemplate_CheckedChanged(object sender, EventArgs e)
        {
            cbBibleNotebook.Enabled = !chkCreateBibleNotebookFromTemplate.Checked;
            btnBibleNotebookSetPath.Enabled = chkCreateBibleNotebookFromTemplate.Checked;
        }

        private void chkCreateBibleCommentsNotebookFromTemplate_CheckedChanged(object sender, EventArgs e)
        {
            cbBibleCommentsNotebook.Enabled = !chkCreateBibleCommentsNotebookFromTemplate.Checked;
            btnBibleCommentsNotebookSetPath.Enabled = chkCreateBibleCommentsNotebookFromTemplate.Checked;
        }

        private void chkCreateBibleStudyNotebookFromTemplate_CheckedChanged(object sender, EventArgs e)
        {
            cbBibleStudyNotebook.Enabled = !chkCreateBibleStudyNotebookFromTemplate.Checked;
            btnBibleStudyNotebookSetPath.Enabled = chkCreateBibleStudyNotebookFromTemplate.Checked;
        }

        private void btnSingleNotebookParameters_Click(object sender, EventArgs e)
        {   
            if (!string.IsNullOrEmpty((string)cbSingleNotebook.SelectedItem))
            {
                string notebookName = (string)cbSingleNotebook.SelectedItem;
                string notebookId = OneNoteUtils.GetNotebookIdByName(_oneNoteApp, notebookName);
                if (NotebookChecker.CheckNotebook(_oneNoteApp, notebookId, NotebookType.Single))
                {
                    if (_notebookParametersForm == null)
                        _notebookParametersForm = new NotebookParametersForm(_oneNoteApp, notebookId);

                    if (_notebookParametersForm.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {   
                        SettingsManager.Instance.SectionGroupId_Bible = _notebookParametersForm.GroupedSectionGroups[SectionGroupType.Bible];                        
                        SettingsManager.Instance.SectionGroupId_BibleComments = _notebookParametersForm.GroupedSectionGroups[SectionGroupType.BibleComments];                        
                        SettingsManager.Instance.SectionGroupId_BibleStudy = _notebookParametersForm.GroupedSectionGroups[SectionGroupType.BibleStudy];

                        _wasSearchedSectionGroupsInSingleNotebook = true;  // нашли необходимые группы секций. 
                    }
                }
                else
                {
                    Logger.LogError(string.Format("Указана неподходящая записная книжка '{0}'.", notebookName));                    
                }
            }
            else
            {
                Logger.LogMessage("Не указана записная книжка.");
            }
        }

        private void btnSingleNotebookSetPath_Click(object sender, EventArgs e)
        {
            if (chkCreateSingleNotebookFromTemplate.Checked)
            {
                if (folderBrowserDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    SingleNotebookFromTemplatePath = folderBrowserDialog.SelectedPath;
                }                
            }
        }

        private void btnBibleNotebookSetPath_Click(object sender, EventArgs e)
        {
            if (chkCreateBibleNotebookFromTemplate.Checked)
            {
                if (folderBrowserDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    BibleNotebookFromTemplatePath = folderBrowserDialog.SelectedPath;
                }                
            }
        }

        private void btnBibleCommentsNotebookSetPath_Click(object sender, EventArgs e)
        {
            if (chkCreateBibleCommentsNotebookFromTemplate.Checked)
            {
                if (folderBrowserDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    BibleCommentsNotebookFromTemplatePath = folderBrowserDialog.SelectedPath;
                }
            }
        }

        private void btnBibleStudyNotebookSetPath_Click(object sender, EventArgs e)
        {
            if (chkCreateBibleStudyNotebookFromTemplate.Checked)
            {
                if (folderBrowserDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    BibleStudyNotebookFromTemplatePath = folderBrowserDialog.SelectedPath;
                }
            }
        }

        private void chkDefaultPageNameParameters_CheckedChanged(object sender, EventArgs e)
        {
            tbPageDescriptionName.Enabled = !chkDefaultPageNameParameters.Checked;
            tbNotesPageName.Enabled = !chkDefaultPageNameParameters.Checked;
            tbBookOverviewName.Enabled = !chkDefaultPageNameParameters.Checked;
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            StopExternalProcess = true;            
        }

        private void MainForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            BibleCommon.Services.Logger.Done();        
        }

        private void btnRelinkComments_Click(object sender, EventArgs e)
        {
            new RelinkAllBibleCommentsManager(_oneNoteApp, this).RelinkAllBibleComments();
        }

        public void PrepareForExternalProcessing(int pbMaxValue, int pbStep, string infoText)
        {
            pbMain.Value = 0;
            pbMain.Maximum = pbMaxValue;
            pbMain.Step = pbStep;
            pbMain.Visible = true;

            tbcMain.Enabled = false;
            lblProgressInfo.Text = infoText;

            btnOK.Enabled = false;
        }

        public void ExternalProcessingDone(string infoText)
        {
            pbMain.Value = 0;
            pbMain.Maximum = 100;
            pbMain.Step = 1;
            pbMain.Visible = false;

            tbcMain.Enabled = true;
            lblProgressInfo.Text = infoText;

            btnOK.Enabled = true;
        }

        public void PerformProgressStep(string infoText)
        {
            lblProgressInfo.Text = infoText;
            System.Windows.Forms.Application.DoEvents();
        }       
    }
}
