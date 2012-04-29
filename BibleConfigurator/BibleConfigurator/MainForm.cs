using System;
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
using BibleCommon.Consts;
using System.Runtime.InteropServices;

namespace BibleConfigurator
{
    public partial class MainForm : Form
    {
        internal class ComboBoxItem
        {
            public string Value { get; set; }
            public object Key { get; set; }

            public override string ToString()
            {
                return Value;
            }
        }

        private Microsoft.Office.Interop.OneNote.Application _oneNoteApp = new Microsoft.Office.Interop.OneNote.Application();

        private string SingleNotebookFromTemplatePath { get; set; }
        private string BibleNotebookFromTemplatePath { get; set; }
        private string BibleCommentsNotebookFromTemplatePath { get; set; }
        private string BibleNotesPagesNotebookFromTemplatePath { get; set; }
        private string BibleStudyNotebookFromTemplatePath { get; set; }

        private bool _wasSearchedSectionGroupsInSingleNotebook = false;       
        

        private const int LoadParametersAttemptsCount = 40;         // количество попыток загрузки параметров после команды создания записных книжек из шаблона
        private const int LoadParametersPauseBetweenAttempts = 10;             // количество секунд ожидания между попытками загрузки параметров
        private const string LoadParametersImageFileName = "loader.gif";


        private NotebookParametersForm _notebookParametersForm = null;
        private bool _runAfterSetup = false;

        public MainForm(params string[] args)
        {
            if (args.Contains("-RunAfterSetup"))
                _runAfterSetup = true;


            // настройки не влияют на работу в дизайнере
            if (this.Site == null || !this.Site.DesignMode)
            {
                // устанавливаем культуру обязательно до InitializeComponent();
                Thread.CurrentThread.CurrentUICulture = LanguageManager.UserLanguage;              
            }


            InitializeComponent();
            BibleCommon.Services.Logger.Init("BibleConfigurator");
        }

        public bool StopExternalProcess { get; set; }
        

        private void btnOK_Click(object sender, EventArgs e)
        {          
            btnOK.Enabled = false;
            lblWarning.Text = string.Empty;
            try
            {
                Logger.Initialize();               

                if (rbSingleNotebook.Checked)
                {
                    SaveSingleNotebookParameters();
                }
                else
                {
                    SettingsManager.Instance.SectionGroupId_Bible = string.Empty;
                    SettingsManager.Instance.SectionGroupId_BibleStudy = string.Empty;
                    SettingsManager.Instance.SectionGroupId_BibleComments = string.Empty;
                    SettingsManager.Instance.SectionGroupId_BibleNotesPages = string.Empty;                    

                    SaveBibleNotebookParameters();             

                    if (!Logger.WasErrorLogged)
                    {
                        SaveBibleStudyNotebookParameters();              

                        if (!Logger.WasErrorLogged)
                        {
                            SaveBibleCommentsNotebookParameters();

                            if (!Logger.WasErrorLogged)
                            {
                                SaveBibleNotesPagesNotebookParameters();
                            }
                        }
                    }
                }

                if (!Logger.WasErrorLogged)
                {
                    SetProgramParameters();

                    SettingsManager.Instance.Save();
                    Close();
                }

            }
            catch (SaveParametersException ex)
            {
                Logger.LogError(ex.Message);
                if (ex.NeedToReload)
                    LoadParameters();                
            }
            finally
            {
                btnOK.Enabled = true;
            }
        }

        private void SaveBibleNotesPagesNotebookParameters()
        {
            string notebookId;
            string notebookName;

            if (chkCreateBibleNotesPagesNotebookFromTemplate.Checked)
            {
                notebookName = CreateNotebookFromTemplate(Consts.BibleNotesPagesNotebookTemplateFileName, BibleNotesPagesNotebookFromTemplatePath);
                if (!string.IsNullOrEmpty(notebookName))
                    WaitAndLoadParameters(NotebookType.BibleNotesPages, notebookName);                         // выйдем из метода только когда OneNote отработает                                                
            }
            else
            {
                notebookName = (string)cbBibleNotesPagesNotebook.SelectedItem;
                TryToLoadNotebookParameters(NotebookType.BibleNotesPages, notebookName, out notebookId);
            }
        }

        private void SaveBibleCommentsNotebookParameters()
        {
            string notebookId;
            string notebookName;

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
        }

        private void SaveBibleStudyNotebookParameters()
        {
            string notebookId;
            string notebookName;

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
        }

        private void SaveBibleNotebookParameters()
        {
            string notebookId;
            string notebookName;

            if (chkCreateBibleNotebookFromTemplate.Checked)
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

        private void SaveSingleNotebookParameters()
        {
            string notebookId;
            string notebookName;

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
                            Logger.LogError(BibleCommon.Resources.Constants.ConfiguratorWrongNotebookSelected);
                        }
                    }
                }

            }
        }


        private void SetProgramParameters()
        {
            if (chkDefaultPageNameParameters.Checked)
            {
                SettingsManager.Instance.LoadDefaultSettings();                
            }
            else
            {
                if (!string.IsNullOrEmpty(tbBookOverviewName.Text))
                    SettingsManager.Instance.PageName_DefaultBookOverview = tbBookOverviewName.Text;

                if (!string.IsNullOrEmpty(tbNotesPageName.Text))
                    SettingsManager.Instance.PageName_Notes = tbNotesPageName.Text;

                if (!string.IsNullOrEmpty(tbPageDescriptionName.Text))
                    SettingsManager.Instance.PageName_DefaultComments = tbPageDescriptionName.Text;

                if (!string.IsNullOrEmpty(tbNotesPageWidth.Text))
                {
                    int notesPageWidth;
                    if (!int.TryParse(tbNotesPageWidth.Text, out notesPageWidth) || notesPageWidth < 200 || notesPageWidth > 1000)
                        throw new SaveParametersException(string.Format("{0} '{1}'", BibleCommon.Resources.Constants.ConfiguratorWrongParameterValue, lblNotesPageWidth.Text), false);

                    SettingsManager.Instance.PageWidth_Notes = notesPageWidth;
                }

                if (!string.IsNullOrEmpty(tbRubbishNotesPageWidth.Text))
                {
                    int rubbishNotesPageWidth;
                    if (!int.TryParse(tbRubbishNotesPageWidth.Text, out rubbishNotesPageWidth) || rubbishNotesPageWidth < 200 || rubbishNotesPageWidth > 1000)
                        throw new SaveParametersException(string.Format("{0} '{1}'", BibleCommon.Resources.Constants.ConfiguratorWrongParameterValue, lblRubbishNotesPageWidth.Text), false);
                    SettingsManager.Instance.PageWidth_RubbishNotes = rubbishNotesPageWidth;
                }

                SettingsManager.Instance.ExpandMultiVersesLinking = chkExpandMultiVersesLinking.Checked;
                SettingsManager.Instance.ExcludedVersesLinking = chkExcludedVersesLinking.Checked;
                SettingsManager.Instance.UseDifferentPagesForEachVerse = chkUseDifferentPages.Checked;

                SettingsManager.Instance.RubbishPage_Use = chkUseRubbishPage.Checked;
                SettingsManager.Instance.PageName_RubbishNotes = tbRubbishNotesPageName.Text;
                SettingsManager.Instance.RubbishPage_ExpandMultiVersesLinking = chkRubbishExpandMultiVersesLinking.Checked;
                SettingsManager.Instance.RubbishPage_ExcludedVersesLinking = chkRubbishExcludedVersesLinking.Checked;



                SettingsManager.Instance.Language = (int)((ComboBoxItem)cbLanguage.SelectedItem).Key;
            }
        }

        private void WaitAndLoadParameters(NotebookType notebookType, string notebookName)
        {   
            PrepareForExternalProcessing(100, 1, string.Format("{0} '{1}'", BibleCommon.Resources.Constants.ConfiguratorNotebookCreation, notebookName));
            
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
                throw new SaveParametersException(BibleCommon.Resources.Constants.ConfiguratorCanNotRequestDataFromOneNote, true);
        }

        private bool TryToLoadNotebookParameters(NotebookType notebookType, string notebookName, out string notebookId, bool silientMode = false)
        {
            notebookId = string.Empty;

            try
            {
                notebookId = OneNoteUtils.GetNotebookIdByName(_oneNoteApp, notebookName, true);                
                if (NotebookChecker.CheckNotebook(_oneNoteApp, notebookId, notebookType))
                {
                    switch (notebookType)
                    {
                        case NotebookType.Single:
                            SettingsManager.Instance.NotebookId_Bible = notebookId;
                            SettingsManager.Instance.NotebookId_BibleComments = notebookId;
                            SettingsManager.Instance.NotebookId_BibleNotesPages = notebookId;
                            SettingsManager.Instance.NotebookId_BibleStudy = notebookId;
                            break;
                        case NotebookType.Bible:
                            SettingsManager.Instance.NotebookId_Bible = notebookId;
                            break;
                        case NotebookType.BibleComments:
                            SettingsManager.Instance.NotebookId_BibleComments = notebookId;
                            break;
                        case NotebookType.BibleNotesPages:
                            SettingsManager.Instance.NotebookId_BibleNotesPages = notebookId;
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
                        Logger.LogError(string.Format("{0} '{1}'.", BibleCommon.Resources.Constants.ConfiguratorWrongNotebookSelected, notebookName));
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
            OneNoteProxy.HierarchyElement notebook = OneNoteProxy.Instance.GetHierarchy(_oneNoteApp, notebookId, HierarchyScope.hsSections, true);

            List<SectionGroupType> sectionGroups = new List<SectionGroupType>();

            foreach (XElement sectionGroup in notebook.Content.Root.XPathSelectElements("one:SectionGroup", notebook.Xnm).Where(sg => !OneNoteUtils.IsRecycleBin(sg)))
            {
                string id = (string)sectionGroup.Attribute("ID");

                if (NotebookChecker.ElementIsBible(sectionGroup, notebook.Xnm) && !sectionGroups.Contains(SectionGroupType.Bible))
                {
                    SettingsManager.Instance.SectionGroupId_Bible = id;
                    sectionGroups.Add(SectionGroupType.Bible);
                }
                else if (NotebookChecker.ElementIsBibleComments(sectionGroup, notebook.Xnm) && !sectionGroups.Contains(SectionGroupType.BibleComments))
                {
                    SettingsManager.Instance.SectionGroupId_BibleComments = id;
                    SettingsManager.Instance.SectionGroupId_BibleNotesPages = id;
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
            OneNoteProxy.HierarchyElement notebook = OneNoteProxy.Instance.GetHierarchy(_oneNoteApp, notebookId, HierarchyScope.hsSections, true);     

            foreach (string sectionGroupId in renamedSectionGroups.Keys)
            {
                XElement sectionGroup = notebook.Content.Root.XPathSelectElement(string.Format("one:SectionGroup[@ID='{0}']", sectionGroupId), notebook.Xnm);

                if (sectionGroup != null)
                {
                    sectionGroup.SetAttributeValue("name", renamedSectionGroups[sectionGroupId]);
                }
                else
                    Logger.LogError(string.Format("{0} '{1}'.", BibleCommon.Resources.Constants.ConfiguratorSectionGroupNotFound, sectionGroupId));
            }

            _oneNoteApp.UpdateHierarchy(notebook.Content.ToString());
            OneNoteProxy.Instance.RefreshHierarchyCache(_oneNoteApp, notebookId, HierarchyScope.hsSections);     
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
                        Logger.LogError(string.Format("{0} '{1}'.", BibleCommon.Resources.Constants.ConfiguratorErrorWhileNotebookOpenning, notebookTemplateFileName));

                    return Path.GetFileNameWithoutExtension(folderPath);
                }
                else
                    Logger.LogError(BibleCommon.Resources.Constants.ConfiguratorSelectAnotherFolder);
            }
            else
                Logger.LogError(string.Format("{0} '{1}'.", BibleCommon.Resources.Constants.ConfiguratorNotebookTemplateNotFound, packageFilePath));

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

        private LoadForm _loadForm;
        private bool _firstShown = true;
        private void MainForm_Shown(object sender, EventArgs e)
        {
            if (_firstShown)
            {
                try
                {
                    if (_runAfterSetup)
                    {
                        if (SettingsManager.Instance.IsConfigured(_oneNoteApp))
                        {
                            this.Close();
                            return;
                        }
                    }
                    PrepareFolderBrowser();
                    SetNotebooksDefaultPaths();

                    LoadParameters();

                    this.Text += string.Format(" v{0}", SettingsManager.Instance.CurrentVersion);
                    this.SetFocus();
                    _firstShown = false;
                }
                finally
                {
                    _loadForm.Close();
                }
            }
        }
      
      
        private void MainForm_Load(object sender, EventArgs e)
        {
            _loadForm = new LoadForm();
            _loadForm.Show();           
        }

        private void LoadParameters()
        {
            if (!SettingsManager.Instance.IsConfigured(_oneNoteApp))
                lblWarning.Text = BibleCommon.Resources.Constants.ConfiguratorNeedSaveChanges;

            Dictionary<string, string> notebooks = GetNotebooks();
            string singleNotebookId = SearchForNotebook(notebooks.Keys, NotebookType.Single);
            string bibleNotebookId = SearchForNotebook(notebooks.Keys, NotebookType.Bible);
            string bibleCommentsNotebookId = SearchForNotebook(notebooks.Keys, NotebookType.BibleComments);            
            string bibleStudyNotebookId = SearchForNotebook(notebooks.Keys, NotebookType.BibleStudy);
            string bibleNotesPagesNotebookId = SearchForNotebook(notebooks.Keys.ToList().Where(s => s != bibleCommentsNotebookId), NotebookType.BibleNotesPages);
            if (string.IsNullOrEmpty(bibleNotesPagesNotebookId))
                bibleNotesPagesNotebookId = bibleCommentsNotebookId;

            rbSingleNotebook.Checked = SettingsManager.Instance.IsSingleNotebook 
                                    && !string.IsNullOrEmpty(singleNotebookId);

            rbMultiNotebook.Checked = !rbSingleNotebook.Checked;
            rbMultiNotebook_CheckedChanged(this, null);            

            cbSingleNotebook.DataSource = notebooks.Values.ToList();
            cbBibleNotebook.DataSource = notebooks.Values.ToList();
            cbBibleCommentsNotebook.DataSource = notebooks.Values.ToList();
            cbBibleNotesPagesNotebook.DataSource = notebooks.Values.ToList();
            cbBibleStudyNotebook.DataSource = notebooks.Values.ToList();

            SetNotebookParameters(rbSingleNotebook.Checked, !string.IsNullOrEmpty(singleNotebookId) ? notebooks[singleNotebookId] :  Consts.SingleNotebookDefaultName, 
                notebooks, SettingsManager.Instance.NotebookId_Bible, cbSingleNotebook, chkCreateSingleNotebookFromTemplate);

            SetNotebookParameters(rbMultiNotebook.Checked, !string.IsNullOrEmpty(bibleNotebookId) ? notebooks[bibleNotebookId] :  Consts.BibleNotebookDefaultName, 
                notebooks, SettingsManager.Instance.NotebookId_Bible, cbBibleNotebook, chkCreateBibleNotebookFromTemplate);

            SetNotebookParameters(rbMultiNotebook.Checked, !string.IsNullOrEmpty(bibleCommentsNotebookId) ? notebooks[bibleCommentsNotebookId] :  Consts.BibleCommentsNotebookDefaultName, 
                notebooks, SettingsManager.Instance.NotebookId_BibleComments, cbBibleCommentsNotebook, chkCreateBibleCommentsNotebookFromTemplate);

            SetNotebookParameters(rbMultiNotebook.Checked, !string.IsNullOrEmpty(bibleNotesPagesNotebookId) ? notebooks[bibleNotesPagesNotebookId] : Consts.BibleNotesPagesNotebookDefaultName,
                notebooks, SettingsManager.Instance.NotebookId_BibleNotesPages, cbBibleNotesPagesNotebook, chkCreateBibleNotesPagesNotebookFromTemplate);

            SetNotebookParameters(rbMultiNotebook.Checked, !string.IsNullOrEmpty(bibleStudyNotebookId) ? notebooks[bibleStudyNotebookId] :  Consts.BibleStudyNotebookDefaultName, 
                notebooks, SettingsManager.Instance.NotebookId_BibleStudy, cbBibleStudyNotebook, chkCreateBibleStudyNotebookFromTemplate);

            tbBookOverviewName.Text = SettingsManager.Instance.PageName_DefaultBookOverview;
            tbNotesPageName.Text = SettingsManager.Instance.PageName_Notes;
            tbPageDescriptionName.Text = SettingsManager.Instance.PageName_DefaultComments;
            tbNotesPageWidth.Text = SettingsManager.Instance.PageWidth_Notes.ToString();
            chkExpandMultiVersesLinking.Checked = SettingsManager.Instance.ExpandMultiVersesLinking;
            chkExcludedVersesLinking.Checked = SettingsManager.Instance.ExcludedVersesLinking;
            chkUseDifferentPages.Checked = SettingsManager.Instance.UseDifferentPagesForEachVerse;

            chkUseRubbishPage.Checked = SettingsManager.Instance.RubbishPage_Use;
            tbRubbishNotesPageName.Text = SettingsManager.Instance.PageName_RubbishNotes;
            tbRubbishNotesPageWidth.Text = SettingsManager.Instance.PageWidth_RubbishNotes.ToString();
            chkRubbishExpandMultiVersesLinking.Checked = SettingsManager.Instance.RubbishPage_ExpandMultiVersesLinking;
            chkRubbishExcludedVersesLinking.Checked = SettingsManager.Instance.RubbishPage_ExcludedVersesLinking;

            chkUseRubbishPage_CheckedChanged(this, new EventArgs());

            InitLanguagesMenu();
        }

        private void InitLanguagesMenu()
        {
            var languages = LanguageManager.GetDisplayedNames();

            var currentLanguage = LanguageManager.UserLanguage;

            foreach (var pair in languages)
            {
                cbLanguage.Items.Add(new ComboBoxItem() { Key = pair.Key, Value = pair.Value });
                if (pair.Key == currentLanguage.LCID)
                    cbLanguage.SelectedIndex = cbLanguage.Items.Count - 1;

            }
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

        private static void SetNotebookParameters(bool loadNameFromSettings, string defaultName, Dictionary<string, string> notebooks, 
            string notebookIdFromSettings, ComboBox cb, CheckBox chk)
        {
            string notebookName = (loadNameFromSettings && !string.IsNullOrEmpty(notebookIdFromSettings)) ? TryToGetNotebookName(notebooks, notebookIdFromSettings) : string.Empty;
            if (!string.IsNullOrEmpty(notebookName) && cb.Items.Contains(notebookName))
                cb.SelectedItem = notebookName;
            else if (cb.Items.Contains(defaultName))
                cb.SelectedItem = defaultName;
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
            BibleNotesPagesNotebookFromTemplatePath = folderBrowserDialog.SelectedPath;
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

            folderBrowserDialog.Description = BibleCommon.Resources.Constants.ConfiguratorSetNotebookFolder;
            folderBrowserDialog.ShowNewFolderButton = true;            
        }

        public Dictionary<string, string> GetNotebooks()
        {
            Dictionary<string, string> result = new Dictionary<string, string>();
            
            OneNoteProxy.HierarchyElement hierarchy = OneNoteProxy.Instance.GetHierarchy(_oneNoteApp, null, HierarchyScope.hsNotebooks, true);

            foreach (XElement notebook in hierarchy.Content.Root.XPathSelectElements("one:Notebook", hierarchy.Xnm))
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
            cbBibleNotesPagesNotebook.Enabled = rbMultiNotebook.Checked;
            cbBibleStudyNotebook.Enabled = rbMultiNotebook.Checked;
            chkCreateBibleNotebookFromTemplate.Enabled = rbMultiNotebook.Checked;
            chkCreateBibleCommentsNotebookFromTemplate.Enabled = rbMultiNotebook.Checked;
            chkCreateBibleNotesPagesNotebookFromTemplate.Enabled = rbMultiNotebook.Checked;
            chkCreateBibleStudyNotebookFromTemplate.Enabled = rbMultiNotebook.Checked;
            btnBibleNotebookSetPath.Enabled = rbMultiNotebook.Checked;
            btnBibleCommentsNotebookSetPath.Enabled = rbMultiNotebook.Checked;
            btnBibleNotesPagesNotebookSetPath.Enabled = rbMultiNotebook.Checked;
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
                chkCreateBibleNotesPagesNotebookFromTemplate_CheckedChanged(this, null);
            }            
        }

        private void chkCreateSingleNotebookFromTemplate_CheckedChanged(object sender, EventArgs e)
        {
            cbSingleNotebook.Enabled = chkCreateSingleNotebookFromTemplate.Enabled && !chkCreateSingleNotebookFromTemplate.Checked;
            btnSingleNotebookParameters.Enabled = chkCreateSingleNotebookFromTemplate.Enabled && !chkCreateSingleNotebookFromTemplate.Checked;
            btnSingleNotebookSetPath.Enabled = chkCreateSingleNotebookFromTemplate.Enabled && chkCreateSingleNotebookFromTemplate.Checked;
        }

        private void chkCreateBibleNotebookFromTemplate_CheckedChanged(object sender, EventArgs e)
        {
            cbBibleNotebook.Enabled = chkCreateBibleNotebookFromTemplate.Enabled && !chkCreateBibleNotebookFromTemplate.Checked;
            btnBibleNotebookSetPath.Enabled = chkCreateBibleNotebookFromTemplate.Enabled && chkCreateBibleNotebookFromTemplate.Checked;
        }

        private void chkCreateBibleCommentsNotebookFromTemplate_CheckedChanged(object sender, EventArgs e)
        {
            cbBibleCommentsNotebook.Enabled = chkCreateBibleCommentsNotebookFromTemplate.Enabled && !chkCreateBibleCommentsNotebookFromTemplate.Checked;
            btnBibleCommentsNotebookSetPath.Enabled = chkCreateBibleCommentsNotebookFromTemplate.Enabled && chkCreateBibleCommentsNotebookFromTemplate.Checked;
        }

        private void chkCreateBibleNotesPagesNotebookFromTemplate_CheckedChanged(object sender, EventArgs e)
        {
            cbBibleNotesPagesNotebook.Enabled = chkCreateBibleNotesPagesNotebookFromTemplate.Enabled && !chkCreateBibleNotesPagesNotebookFromTemplate.Checked;
            btnBibleNotesPagesNotebookSetPath.Enabled = chkCreateBibleNotesPagesNotebookFromTemplate.Enabled && chkCreateBibleNotesPagesNotebookFromTemplate.Checked;
        }

        private void chkCreateBibleStudyNotebookFromTemplate_CheckedChanged(object sender, EventArgs e)
        {
            cbBibleStudyNotebook.Enabled = chkCreateBibleStudyNotebookFromTemplate.Enabled && !chkCreateBibleStudyNotebookFromTemplate.Checked;
            btnBibleStudyNotebookSetPath.Enabled = chkCreateBibleStudyNotebookFromTemplate.Enabled && chkCreateBibleStudyNotebookFromTemplate.Checked;
        }

        private void btnSingleNotebookParameters_Click(object sender, EventArgs e)
        {   
            if (!string.IsNullOrEmpty((string)cbSingleNotebook.SelectedItem))
            {
                string notebookName = (string)cbSingleNotebook.SelectedItem;
                string notebookId = OneNoteUtils.GetNotebookIdByName(_oneNoteApp, notebookName, true);
                if (NotebookChecker.CheckNotebook(_oneNoteApp, notebookId, NotebookType.Single))
                {
                    if (_notebookParametersForm == null)
                        _notebookParametersForm = new NotebookParametersForm(_oneNoteApp, notebookId);

                    if (_notebookParametersForm.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {   
                        SettingsManager.Instance.SectionGroupId_Bible = _notebookParametersForm.GroupedSectionGroups[SectionGroupType.Bible];
                        SettingsManager.Instance.SectionGroupId_BibleStudy = _notebookParametersForm.GroupedSectionGroups[SectionGroupType.BibleStudy];
                        SettingsManager.Instance.SectionGroupId_BibleComments = _notebookParametersForm.GroupedSectionGroups[SectionGroupType.BibleComments];
                        SettingsManager.Instance.SectionGroupId_BibleNotesPages = _notebookParametersForm.GroupedSectionGroups[SectionGroupType.BibleComments];

                        _wasSearchedSectionGroupsInSingleNotebook = true;  // нашли необходимые группы секций. 
                    }
                }
                else
                {
                    Logger.LogError(string.Format("{0} '{1}'.", BibleCommon.Resources.Constants.ConfiguratorWrongNotebookSelected, notebookName));                    
                }
            }
            else
            {
                Logger.LogMessage(BibleCommon.Resources.Constants.ConfiguratorNotebookNotDefined);
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

        private void btnBibleNotesPagesNotebookSetPath_Click(object sender, EventArgs e)
        {

            if (chkCreateBibleNotesPagesNotebookFromTemplate.Checked)
            {
                if (folderBrowserDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    BibleNotesPagesNotebookFromTemplatePath = folderBrowserDialog.SelectedPath;
                }
            }
        }

        private void chkDefaultPageNameParameters_CheckedChanged(object sender, EventArgs e)
        {
            tbPageDescriptionName.Enabled = !chkDefaultPageNameParameters.Checked;
            tbNotesPageName.Enabled = !chkDefaultPageNameParameters.Checked;
            tbBookOverviewName.Enabled = !chkDefaultPageNameParameters.Checked;
            tbNotesPageWidth.Enabled = !chkDefaultPageNameParameters.Checked;
            chkExpandMultiVersesLinking.Enabled = !chkDefaultPageNameParameters.Checked;
            chkExcludedVersesLinking.Enabled = !chkDefaultPageNameParameters.Checked;
            chkUseDifferentPages.Enabled = !chkDefaultPageNameParameters.Checked;
            chkUseRubbishPage.Enabled = !chkDefaultPageNameParameters.Checked;
            tbRubbishNotesPageName.Enabled = !chkDefaultPageNameParameters.Checked;
            tbRubbishNotesPageWidth.Enabled = !chkDefaultPageNameParameters.Checked;
            chkRubbishExpandMultiVersesLinking.Enabled = !chkDefaultPageNameParameters.Checked;
            chkRubbishExcludedVersesLinking.Enabled = !chkDefaultPageNameParameters.Checked;

            chkUseRubbishPage_CheckedChanged(this, new EventArgs());            
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
            FormExtensions.SetControlPropertyThreadSafe(pbMain, "Value", 0);
            FormExtensions.SetControlPropertyThreadSafe(pbMain, "Maximum", 100);
            FormExtensions.SetControlPropertyThreadSafe(pbMain, "Step", 1);
            FormExtensions.SetControlPropertyThreadSafe(pbMain, "Visible", false);
            FormExtensions.SetControlPropertyThreadSafe(tbcMain, "Enabled", true);
            FormExtensions.SetControlPropertyThreadSafe(lblProgressInfo, "Text", infoText);
            FormExtensions.SetControlPropertyThreadSafe(btnOK, "Enabled", true);

            //pbMain.Value = 0;
            //pbMain.Maximum = 100;
            //pbMain.Step = 1;
            //pbMain.Visible = false;

            //tbcMain.Enabled = true;
            //lblProgressInfo.Text = infoText;

            //btnOK.Enabled = true;
        }

        public void PerformProgressStep(string infoText)
        {
            FormExtensions.SetControlPropertyThreadSafe(lblProgressInfo, "Text", infoText);
            //lblProgressInfo.Text = infoText;            
            if (pbMain.Value < pbMain.Maximum)
                FormExtensions.SetControlPropertyThreadSafe(pbMain, "Value", pbMain.Value + 1);                
            //pbMain.PerformStep();            
            System.Windows.Forms.Application.DoEvents();
        }

        private void chkUseRubbishPage_CheckedChanged(object sender, EventArgs e)
        {
            tbRubbishNotesPageName.Enabled = chkUseRubbishPage.Enabled && chkUseRubbishPage.Checked;
            tbRubbishNotesPageWidth.Enabled = chkUseRubbishPage.Enabled && chkUseRubbishPage.Checked;
            chkRubbishExpandMultiVersesLinking.Enabled = chkUseRubbishPage.Enabled && chkUseRubbishPage.Checked;
            chkRubbishExcludedVersesLinking.Enabled = chkUseRubbishPage.Enabled && chkUseRubbishPage.Checked;            
        }

        private void btnDeleteNotesPages_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show(BibleCommon.Resources.Constants.ConfiguratorQuestionDeleteAllNotesPages, BibleCommon.Resources.Constants.Warning,
                        MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) == System.Windows.Forms.DialogResult.Yes)
            {
                if (SettingsManager.Instance.NotebookId_BibleComments == SettingsManager.Instance.NotebookId_BibleNotesPages
                    || MessageBox.Show(BibleCommon.Resources.Constants.ConfiguratorQuestionDeleteAllNotesPagesManually, BibleCommon.Resources.Constants.Warning,
                        MessageBoxButtons.YesNo, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2) == System.Windows.Forms.DialogResult.Yes)
                    new DeleteNotesPagesManager(_oneNoteApp, this).DeleteNotesPages();
            }
        }

        private void btnResizeBibleTables_Click(object sender, EventArgs e)
        {
            SetWidthForm form = new SetWidthForm();
            if (form.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                ResizeBibleManager manager = new ResizeBibleManager(_oneNoteApp, this);
                manager.ResizeBiblePages(form.BiblePagesWidth);
            }
        }

        private void btnBackup_Click(object sender, EventArgs e)
        {
            saveFileDialog.DefaultExt = ".zip";
            saveFileDialog.FileName = string.Format("{0}_backup_{1}", Constants.ToolsName, DateTime.Now.ToShortDateString());
            
            if (saveFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                BackupManager manager = new BackupManager(_oneNoteApp, this);
                manager.Backup(saveFileDialog.FileName);
            }
        }

      
    }
}
