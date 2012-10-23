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
using BibleCommon.Common;
using BibleConfigurator.ModuleConverter;
using System.Globalization;

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
        

        private const int LoadParametersAttemptsCount = 80;         // количество попыток загрузки параметров после команды создания записных книжек из шаблона
        private const int LoadParametersPauseBetweenAttempts = 5;             // количество секунд ожидания между попытками загрузки параметров
        private const string LoadParametersImageFileName = "loader.gif";

        protected CustomFormLogger LongProcessLogger { get; set; }

        private NotebookParametersForm _notebookParametersForm = null;
        
        public bool ShowModulesTabAtStartUp { get; set; }
        public bool NeedToSaveChangesAfterLoadingModuleAtStartUp { get; set; }

        public MainForm(params string[] args)
        {
            this.SetFormUICulture();

            InitializeComponent();
            BibleCommon.Services.Logger.Init("BibleConfigurator");
            LongProcessLogger = new CustomFormLogger(this);
        }

        public bool StopExternalProcess { get; set; }        

        private void btnOK_Click(object sender, EventArgs e)
        {
            try
            {
                CommitChanges(true);
            }
            catch (Exception ex)
            {                
                FormLogger.LogError(ex);
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnApply_Click(object sender, EventArgs e)
        {
            try
            {
                CommitChanges(false);
            }
            catch (Exception ex)
            {                
                FormLogger.LogError(ex);
            }
        }    

        private void CommitChanges(bool closeForm)
        {
            ModuleInfo module = null;

            try
            {
                module = ModulesManager.GetCurrentModuleInfo();
            }
            catch (InvalidModuleException ex)
            {
                FormLogger.LogMessage(ex.Message);
                return;
            }

            btnOK.Enabled = false;
            btnClose.Enabled = false;
            btnApply.Enabled = false;
            bool lblWarningVisibilityBefore = lblWarning.Visible;
            lblWarning.Visible = false;
            this.TopMost = true;

            try
            {
                FormLogger.Initialize();

                if (rbSingleNotebook.Checked && module.UseSingleNotebook())
                {
                    SaveSingleNotebookParameters(module);
                }
                else
                {
                    SettingsManager.Instance.SectionGroupId_Bible = string.Empty;
                    SettingsManager.Instance.SectionGroupId_BibleStudy = string.Empty;
                    SettingsManager.Instance.SectionGroupId_BibleComments = string.Empty;
                    SettingsManager.Instance.SectionGroupId_BibleNotesPages = string.Empty;

                    SaveMultiNotebookParameters(module, ContainerType.Bible,
                        chkCreateBibleNotebookFromTemplate, cbBibleNotebook, BibleNotebookFromTemplatePath);

                    SaveMultiNotebookParameters(module, ContainerType.BibleStudy,
                        chkCreateBibleStudyNotebookFromTemplate, cbBibleStudyNotebook, BibleStudyNotebookFromTemplatePath);

                    SaveMultiNotebookParameters(module, ContainerType.BibleComments,
                        chkCreateBibleCommentsNotebookFromTemplate, cbBibleCommentsNotebook, BibleCommentsNotebookFromTemplatePath);

                    SaveMultiNotebookParameters(module, ContainerType.BibleNotesPages,
                        chkCreateBibleNotesPagesNotebookFromTemplate, cbBibleNotesPagesNotebook, BibleNotesPagesNotebookFromTemplatePath);                    
                }

                if (!BibleVersesLinksCacheManager.CacheIsActive(SettingsManager.Instance.NotebookId_Bible))
                {
                    IndexBible();                    
                }

                if (!FormLogger.WasErrorLogged)
                {
                    SetProgramParameters();

                    SettingsManager.Instance.Save();
                    if (closeForm)
                        Close();
                    else
                        ReLoadParameters(false);
                }
            }
            catch (SaveParametersException ex)
            {
                FormLogger.LogError(ex);
                if (ex.NeedToReload)
                    LoadParameters(module, null);

                lblWarning.Visible = lblWarningVisibilityBefore;
                tbcMain.SelectedTab = tbcMain.TabPages[tabPage1.Name];
            }
            finally
            {
                btnOK.Enabled = true;
                btnClose.Enabled = true;
                btnApply.Enabled = true;
                this.TopMost = false;
            }
        }

        private void IndexBible()
        {
            int chaptersCount = ModulesManager.GetBibleChaptersCount(SettingsManager.Instance.ModuleName, true);
            PrepareForExternalProcessing(chaptersCount, 1, BibleCommon.Resources.Constants.IndexBibleStart);
            LongProcessLogger.Preffix = string.Format("{0}: ", BibleCommon.Resources.Constants.IndexBible);
            BibleVersesLinksCacheManager.GenerateBibleVersesLinks(_oneNoteApp,
                SettingsManager.Instance.NotebookId_Bible, SettingsManager.Instance.SectionGroupId_Bible, LongProcessLogger);
            ExternalProcessingDone(BibleCommon.Resources.Constants.IndexBibleFinish);
        }

        private void SaveMultiNotebookParameters(ModuleInfo module, ContainerType notebookType,
            CheckBox createFromTemplateControl, ComboBox selectedNotebookNameControl, string notebookFromTemplatePath)
        {
            if (createFromTemplateControl.Checked)
            {
                string notebookTemplateFileName = module.GetNotebook(notebookType).Name;
                string notebookName = CreateNotebookFromTemplate(notebookTemplateFileName, notebookFromTemplatePath);
                if (!string.IsNullOrEmpty(notebookName))
                {
                    WaitAndLoadParameters(notebookType, notebookName);                         // выйдем из метода только когда OneNote отработает
                    createFromTemplateControl.Checked = false;  // чтоб если ошибки будут потом, он заново не создавал
                    selectedNotebookNameControl.Items.Add(notebookName);
                    selectedNotebookNameControl.SelectedItem = notebookName;
                }
            }
            else
            {
                string notebookId;
                TryToLoadNotebookParameters(notebookType, (string)selectedNotebookNameControl.SelectedItem, false, out notebookId);
            }
        }

        private void SaveSingleNotebookParameters(ModuleInfo module)
        {
            string notebookId;
            string notebookName;

            if (chkCreateSingleNotebookFromTemplate.Checked)
            {
                string notebookTemplateFileName = module.GetNotebook(ContainerType.Single).Name;
                notebookName = CreateNotebookFromTemplate(notebookTemplateFileName, SingleNotebookFromTemplatePath);
                if (!string.IsNullOrEmpty(notebookName))
                {
                    WaitAndLoadParameters(ContainerType.Single, notebookName);
                    SearchForCorrespondenceSectionGroups(module, SettingsManager.Instance.NotebookId_Bible);
                }
            }
            else
            {
                notebookName = (string)cbSingleNotebook.SelectedItem;
                if (TryToLoadNotebookParameters(ContainerType.Single, notebookName, false, out notebookId))
                {
                    if (_notebookParametersForm != null && _notebookParametersForm.RenamedSectionGroups.Count > 0)
                        RenameSectionGroupsForm(notebookId, _notebookParametersForm.RenamedSectionGroups);

                    if (!_wasSearchedSectionGroupsInSingleNotebook)
                    {
                        try
                        {
                            SearchForCorrespondenceSectionGroups(module, notebookId);
                        }
                        catch (InvalidNotebookException)
                        {
                            FormLogger.LogError(BibleCommon.Resources.Constants.ConfiguratorWrongNotebookSelected);
                        }
                    }
                }

            }
        }
        
        private void SetProgramParameters()
        {
            bool localeWasChanged = false;
            if (SettingsManager.Instance.Language != (int)((ComboBoxItem)cbLanguage.SelectedItem).Key)
            {
                localeWasChanged = true;
                SettingsManager.Instance.Language = (int)((ComboBoxItem)cbLanguage.SelectedItem).Key;                
            }

            if (chkDefaultPageNameParameters.Checked)
            {
                SettingsManager.Instance.UseDefaultSettings = true;                
            }
            else
            {
                if (WasModified())
                    SettingsManager.Instance.UseDefaultSettings = false;

                SaveIntegerSettings();
                SaveBooleanSettings();                
                SaveLocalazibleSettings(localeWasChanged);
            }  
        }

        private void SaveBooleanSettings()
        {
            SettingsManager.Instance.ExpandMultiVersesLinking = chkExpandMultiVersesLinking.Checked;
            SettingsManager.Instance.ExcludedVersesLinking = chkExcludedVersesLinking.Checked;
            SettingsManager.Instance.UseDifferentPagesForEachVerse = chkUseDifferentPages.Checked;
            SettingsManager.Instance.RubbishPage_Use = chkUseRubbishPage.Checked;
            SettingsManager.Instance.RubbishPage_ExpandMultiVersesLinking = chkRubbishExpandMultiVersesLinking.Checked;
            SettingsManager.Instance.RubbishPage_ExcludedVersesLinking = chkRubbishExcludedVersesLinking.Checked;
        }

        private void SaveIntegerSettings()
        {
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
        }

        private void SaveLocalazibleSettings(bool localeWasChanged)
        {
            CultureInfo resourceCulture = new CultureInfo(SettingsManager.Instance.Language);

            if (!string.IsNullOrEmpty(tbBookOverviewName.Text))
            {
                if (SettingsManager.Instance.SectionName_DefaultBookOverview == tbBookOverviewName.Text
                    && SettingsManager.Instance.SectionName_DefaultBookOverview == BibleCommon.Resources.Constants.DefaultPageNameDefaultBookOverview
                    && localeWasChanged)
                    SettingsManager.Instance.SectionName_DefaultBookOverview = BibleCommon.Resources.Constants.ResourceManager
                        .GetString(BibleCommon.Consts.Constants.ResourceName_DefaultPageNameDefaultBookOverview, resourceCulture);
                else
                    SettingsManager.Instance.SectionName_DefaultBookOverview = tbBookOverviewName.Text;
            }            

            if (!string.IsNullOrEmpty(tbCommentsPageName.Text))
            {
                if (SettingsManager.Instance.PageName_DefaultComments == tbCommentsPageName.Text
                    && SettingsManager.Instance.PageName_DefaultComments == BibleCommon.Resources.Constants.DefaultPageNameDefaultComments
                    && localeWasChanged)
                    SettingsManager.Instance.PageName_DefaultComments = BibleCommon.Resources.Constants.ResourceManager
                        .GetString(BibleCommon.Consts.Constants.ResourceName_DefaultPageNameDefaultComments, resourceCulture);
                else
                    SettingsManager.Instance.PageName_DefaultComments = tbCommentsPageName.Text;
            }

            if (!string.IsNullOrEmpty(tbNotesPageName.Text))
            {
                if (SettingsManager.Instance.PageName_Notes == tbNotesPageName.Text
                    && SettingsManager.Instance.PageName_Notes == BibleCommon.Resources.Constants.DefaultPageName_Notes
                    && localeWasChanged)
                    SettingsManager.Instance.PageName_Notes = BibleCommon.Resources.Constants.ResourceManager
                        .GetString(BibleCommon.Consts.Constants.ResourceName_DefaultPageName_Notes, resourceCulture);
                else
                    SettingsManager.Instance.PageName_Notes = tbNotesPageName.Text;
            }

            if (!string.IsNullOrEmpty(tbRubbishNotesPageName.Text))
            {
                if (SettingsManager.Instance.PageName_RubbishNotes == tbRubbishNotesPageName.Text
                    && SettingsManager.Instance.PageName_RubbishNotes == BibleCommon.Resources.Constants.DefaultPageName_RubbishNotes
                    && localeWasChanged)
                    SettingsManager.Instance.PageName_RubbishNotes = BibleCommon.Resources.Constants.ResourceManager
                        .GetString(BibleCommon.Consts.Constants.ResourceName_DefaultPageName_RubbishNotes, resourceCulture);                   
                else
                    SettingsManager.Instance.PageName_RubbishNotes = tbRubbishNotesPageName.Text;
            }
        }

        private bool WasModified()
        {
            return SettingsManager.Instance.SectionName_DefaultBookOverview != tbBookOverviewName.Text
                || SettingsManager.Instance.PageName_Notes != tbNotesPageName.Text
                || SettingsManager.Instance.PageName_DefaultComments != tbCommentsPageName.Text
                || SettingsManager.Instance.ExpandMultiVersesLinking != chkExpandMultiVersesLinking.Checked
                || SettingsManager.Instance.ExcludedVersesLinking != chkExcludedVersesLinking.Checked
                || SettingsManager.Instance.UseDifferentPagesForEachVerse != chkUseDifferentPages.Checked
                || SettingsManager.Instance.RubbishPage_Use != chkUseRubbishPage.Checked
                || SettingsManager.Instance.PageName_RubbishNotes != tbRubbishNotesPageName.Text
                || SettingsManager.Instance.RubbishPage_ExpandMultiVersesLinking != chkRubbishExpandMultiVersesLinking.Checked
                || SettingsManager.Instance.RubbishPage_ExcludedVersesLinking != chkRubbishExcludedVersesLinking.Checked
                || SettingsManager.Instance.PageWidth_Notes.ToString() != tbNotesPageWidth.Text
                || SettingsManager.Instance.PageWidth_RubbishNotes.ToString() != tbRubbishNotesPageWidth.Text;

        }

        private void WaitAndLoadParameters(ContainerType notebookType, string notebookName)
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
                    
                    if (TryToLoadNotebookParameters(notebookType, notebookName, true, out notebookId))
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

        private bool TryToLoadNotebookParameters(ContainerType notebookType, string notebookName, bool silientMode, out string notebookId)
        {
            notebookId = string.Empty;

            try
            {
                notebookId = OneNoteUtils.GetNotebookIdByName(_oneNoteApp, notebookName, true);
                var module = ModulesManager.GetCurrentModuleInfo();

                string errorText;
                if (NotebookChecker.CheckNotebook(_oneNoteApp, module, notebookId, notebookType, out errorText))
                {
                    switch (notebookType)
                    {
                        case ContainerType.Single:
                            SettingsManager.Instance.NotebookId_Bible = notebookId;
                            SettingsManager.Instance.NotebookId_BibleComments = notebookId;
                            SettingsManager.Instance.NotebookId_BibleNotesPages = notebookId;
                            SettingsManager.Instance.NotebookId_BibleStudy = notebookId;
                            break;
                        case ContainerType.Bible:
                            SettingsManager.Instance.NotebookId_Bible = notebookId;
                            break;
                        case ContainerType.BibleComments:
                            SettingsManager.Instance.NotebookId_BibleComments = notebookId;
                            break;
                        case ContainerType.BibleNotesPages:
                            SettingsManager.Instance.NotebookId_BibleNotesPages = notebookId;
                            break;
                        case ContainerType.BibleStudy:
                            SettingsManager.Instance.NotebookId_BibleStudy = notebookId;
                            break;
                    }

                    return true;
                }
                else
                {
                    string message = string.Format(BibleCommon.Resources.Constants.ConfiguratorWrongNotebookSelected + "\n" + errorText, notebookName, notebookType);
                    
                    if (!silientMode)
                        throw new SaveParametersException(message, false);  
                    else
                        BibleCommon.Services.Logger.LogError(message);
                }
            }
            catch (Exception ex)
            {
                if (!silientMode)
                    throw new SaveParametersException(ex.Message, false);
                else
                    BibleCommon.Services.Logger.LogError(ex);
            }

            return false;
        }

        private void SearchForCorrespondenceSectionGroups(ModuleInfo module, string notebookId)
        {
            OneNoteProxy.HierarchyElement notebook = OneNoteProxy.Instance.GetHierarchy(_oneNoteApp, notebookId, HierarchyScope.hsSections, true);

            List<ContainerType> sectionGroups = new List<ContainerType>();

            foreach (XElement sectionGroup in notebook.Content.Root.XPathSelectElements("one:SectionGroup", notebook.Xnm).Where(sg => !OneNoteUtils.IsRecycleBin(sg)))
            {
                string id = (string)sectionGroup.Attribute("ID");

                if (NotebookChecker.ElementIsBible(module, sectionGroup, notebook.Xnm) && !sectionGroups.Contains(ContainerType.Bible))
                {
                    SettingsManager.Instance.SectionGroupId_Bible = id;
                    sectionGroups.Add(ContainerType.Bible);
                }
                else if (NotebookChecker.ElementIsBibleComments(module, sectionGroup, notebook.Xnm) && !sectionGroups.Contains(ContainerType.BibleComments))
                {
                    SettingsManager.Instance.SectionGroupId_BibleComments = id;
                    SettingsManager.Instance.SectionGroupId_BibleNotesPages = id;
                    sectionGroups.Add(ContainerType.BibleComments);
                }
                else if (!sectionGroups.Contains(ContainerType.BibleStudy))
                {
                    SettingsManager.Instance.SectionGroupId_BibleStudy = id;
                    sectionGroups.Add(ContainerType.BibleStudy);
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
                    FormLogger.LogError(string.Format("{0} '{1}'.", BibleCommon.Resources.Constants.ConfiguratorSectionGroupNotFound, sectionGroupId));
            }

            _oneNoteApp.UpdateHierarchy(notebook.Content.ToString(), Constants.CurrentOneNoteSchema);
            OneNoteProxy.Instance.RefreshHierarchyCache(_oneNoteApp, notebookId, HierarchyScope.hsSections);     
        }

        private string CreateNotebookFromTemplate(string notebookTemplateFileName, string notebookFromTemplatePath)
        {
            string s;
            string packageDirectory = ModulesManager.GetCurrentModuleDirectiory();                
            string packageFilePath = Path.Combine(packageDirectory, notebookTemplateFileName);

            if (File.Exists(packageFilePath))
            {
                string folderPath = Path.Combine(notebookFromTemplatePath, Path.GetFileNameWithoutExtension(notebookTemplateFileName));                

                folderPath = Utils.GetNewDirectoryPath(folderPath);

                //if (!string.IsNullOrEmpty(folderPath))
                //{
                    _oneNoteApp.OpenPackage(packageFilePath, folderPath, out s);

                    string[] files = Directory.GetFiles(s, "*.onetoc2", SearchOption.TopDirectoryOnly);
                    if (files.Length > 0)
                        Process.Start(files[0]);
                    else
                        FormLogger.LogError(string.Format("{0} '{1}'.", BibleCommon.Resources.Constants.ConfiguratorErrorWhileNotebookOpenning, notebookTemplateFileName));

                    return Path.GetFileNameWithoutExtension(folderPath);
                //}
                //else
                //    Logger.LogError(BibleCommon.Resources.Constants.ConfiguratorSelectAnotherFolder);
            }
            else
                FormLogger.LogError(string.Format("{0} '{1}'.", BibleCommon.Resources.Constants.ConfiguratorNotebookTemplateNotFound, packageFilePath));

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
                    bool? needSaveSettings = null;

                    if (ShowModulesTabAtStartUp)
                    {                        
                        tbcMain.SelectedTab = tbcMain.TabPages[tabPage4.Name];
                        _wasLoadedModulesInfo = false;                        

                        if (NeedToSaveChangesAfterLoadingModuleAtStartUp)
                            needSaveSettings = true;
                    }
                    
                    PrepareFolderBrowser();
                    SetNotebooksDefaultPaths();

                    if (!SettingsManager.Instance.CurrentModuleIsCorrect())
                        tbcMain.SelectedTab = tbcMain.TabPages[tabPage4.Name];                    
                    else
                    {
                        var module = ModulesManager.GetCurrentModuleInfo();
                        LoadParameters(module, needSaveSettings);
                    }

                    this.Text += string.Format(" v{0}", SettingsManager.Instance.CurrentVersion);
                    this.SetFocus();
                    _firstShown = false;
                }                
                finally
                {
                    _loadForm.Hide();
                }
            }
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            _loadForm = new LoadForm();

            _loadForm.Show();
        }

        private void LoadParameters(ModuleInfo module, bool? needToSaveSettings)
        {
            if (!SettingsManager.Instance.IsConfigured(_oneNoteApp) || needToSaveSettings.GetValueOrDefault(false))
                lblWarning.Visible = true;

            Dictionary<string, string> notebooks = GetNotebooks();
            string singleNotebookId = module.UseSingleNotebook() ? SearchForNotebook(module, notebooks.Keys, ContainerType.Single) : string.Empty;
            string bibleNotebookId = SearchForNotebook(module, notebooks.Keys, ContainerType.Bible);
            string bibleCommentsNotebookId = SearchForNotebook(module, notebooks.Keys, ContainerType.BibleComments);
            string bibleStudyNotebookId = SearchForNotebook(module, notebooks.Keys, ContainerType.BibleStudy);
            string bibleNotesPagesNotebookId = SearchForNotebook(module, notebooks.Keys.ToList().Where(s => s != bibleCommentsNotebookId), ContainerType.BibleNotesPages);
            if (string.IsNullOrEmpty(bibleNotesPagesNotebookId))
                bibleNotesPagesNotebookId = bibleCommentsNotebookId;

            if (string.IsNullOrEmpty(bibleNotebookId) && string.IsNullOrEmpty(bibleCommentsNotebookId))  // а то иначе он всегда "Личную" при установке выбирает
                bibleStudyNotebookId = null;

            rbSingleNotebook.Checked = SettingsManager.Instance.IsSingleNotebook 
                                    && !string.IsNullOrEmpty(singleNotebookId);

            rbMultiNotebook.Checked = !rbSingleNotebook.Checked;
            rbMultiNotebook_CheckedChanged(this, null);

            cbSingleNotebook.Items.Clear();
            cbBibleNotebook.Items.Clear();
            cbBibleCommentsNotebook.Items.Clear();
            cbBibleNotesPagesNotebook.Items.Clear();
            cbBibleStudyNotebook.Items.Clear();

            foreach (var notebook in notebooks.Values)
            {
                cbSingleNotebook.Items.Add(notebook);
                cbBibleNotebook.Items.Add(notebook);
                cbBibleCommentsNotebook.Items.Add(notebook);
                cbBibleNotesPagesNotebook.Items.Add(notebook);
                cbBibleStudyNotebook.Items.Add(notebook);
            }

            if (module.UseSingleNotebook())
            {
                SetNotebookParameters(rbSingleNotebook.Checked, !string.IsNullOrEmpty(singleNotebookId) ? notebooks[singleNotebookId] :
                    Path.GetFileNameWithoutExtension(module.GetNotebook(ContainerType.Single).Name),
                    notebooks, SettingsManager.Instance.NotebookId_Bible, cbSingleNotebook, chkCreateSingleNotebookFromTemplate);
            }
            

            SetNotebookParameters(rbMultiNotebook.Checked, !string.IsNullOrEmpty(bibleNotebookId) ? notebooks[bibleNotebookId] :
                Path.GetFileNameWithoutExtension(module.GetNotebook(ContainerType.Bible).Name), 
                notebooks, SettingsManager.Instance.NotebookId_Bible, cbBibleNotebook, chkCreateBibleNotebookFromTemplate);

            SetNotebookParameters(rbMultiNotebook.Checked, !string.IsNullOrEmpty(bibleStudyNotebookId) ? notebooks[bibleStudyNotebookId] :
                Path.GetFileNameWithoutExtension(module.GetNotebook(ContainerType.BibleStudy).Name),
                notebooks, SettingsManager.Instance.NotebookId_BibleStudy, cbBibleStudyNotebook, chkCreateBibleStudyNotebookFromTemplate);

            SetNotebookParameters(rbMultiNotebook.Checked, !string.IsNullOrEmpty(bibleCommentsNotebookId) ? notebooks[bibleCommentsNotebookId] :
                Path.GetFileNameWithoutExtension(module.GetNotebook(ContainerType.BibleComments).Name), 
                notebooks, SettingsManager.Instance.NotebookId_BibleComments, cbBibleCommentsNotebook, chkCreateBibleCommentsNotebookFromTemplate);

            SetNotebookParameters(rbMultiNotebook.Checked, !string.IsNullOrEmpty(bibleNotesPagesNotebookId) ? notebooks[bibleNotesPagesNotebookId] :
                Path.GetFileNameWithoutExtension(module.GetNotebook(ContainerType.BibleNotesPages).Name), 
                notebooks, SettingsManager.Instance.NotebookId_BibleNotesPages, cbBibleNotesPagesNotebook, chkCreateBibleNotesPagesNotebookFromTemplate);            

            tbBookOverviewName.Text = SettingsManager.Instance.SectionName_DefaultBookOverview;
            tbNotesPageName.Text = SettingsManager.Instance.PageName_Notes;
            tbCommentsPageName.Text = SettingsManager.Instance.PageName_DefaultComments;
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

            if (!rbSingleNotebook.Checked)
                rbSingleNotebook.Enabled = false;
        }

        private void InitLanguagesMenu()
        {
            var languages = LanguageManager.GetDisplayedNames();

            var currentLanguage = LanguageManager.UserLanguage;

            cbLanguage.Items.Clear();
            foreach (var pair in languages)
            {
                cbLanguage.Items.Add(new ComboBoxItem() { Key = pair.Key, Value = pair.Value });
                if (pair.Key == currentLanguage.LCID)
                    cbLanguage.SelectedIndex = cbLanguage.Items.Count - 1;

            }
        }

        private string SearchForNotebook(ModuleInfo module, IEnumerable<string> notebooksIds, ContainerType notebookType)
        {
            foreach (string notebookId in notebooksIds)
            {
                string errorText;
                if (NotebookChecker.CheckNotebook(_oneNoteApp, module, notebookId, notebookType, out errorText))
                {
                    return notebookId;
                }
            }

            return null;
        }

        private static void SetNotebookParameters(bool loadNameFromSettings, string defaultName, Dictionary<string, string> notebooks, 
            string notebookIdFromSettings, ComboBox cb, CheckBox chk)
        {
            chk.Checked = false;
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
            string defaultNotebookFolderPath;
            _oneNoteApp.GetSpecialLocation(SpecialLocation.slDefaultNotebookFolder, out defaultNotebookFolderPath);            
            
            folderBrowserDialog.SelectedPath = defaultNotebookFolderPath;
            folderBrowserDialog.Description = BibleCommon.Resources.Constants.ConfiguratorSetNotebookFolder;
            folderBrowserDialog.ShowNewFolderButton = true;

            string toolTipMessage = BibleCommon.Resources.Constants.DefineNotebookDirectory;
            FormExtensions.SetToolTip(btnSingleNotebookSetPath, toolTipMessage);
            FormExtensions.SetToolTip(btnBibleNotebookSetPath, toolTipMessage);
            FormExtensions.SetToolTip(btnBibleStudyNotebookSetPath, toolTipMessage);
            FormExtensions.SetToolTip(btnBibleCommentsNotebookSetPath, toolTipMessage);
            FormExtensions.SetToolTip(btnBibleNotesPagesNotebookSetPath, toolTipMessage);
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
            lblSelectSingleNotebook.Enabled = rbSingleNotebook.Checked;
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
                var module = ModulesManager.GetCurrentModuleInfo();
                string errorText;
                if (NotebookChecker.CheckNotebook(_oneNoteApp, module, notebookId, ContainerType.Single, out errorText))
                {
                    if (_notebookParametersForm == null)
                        _notebookParametersForm = new NotebookParametersForm(_oneNoteApp, notebookId);

                    if (_notebookParametersForm.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {   
                        SettingsManager.Instance.SectionGroupId_Bible = _notebookParametersForm.GroupedSectionGroups[ContainerType.Bible];
                        SettingsManager.Instance.SectionGroupId_BibleStudy = _notebookParametersForm.GroupedSectionGroups[ContainerType.BibleStudy];
                        SettingsManager.Instance.SectionGroupId_BibleComments = _notebookParametersForm.GroupedSectionGroups[ContainerType.BibleComments];
                        SettingsManager.Instance.SectionGroupId_BibleNotesPages = _notebookParametersForm.GroupedSectionGroups[ContainerType.BibleComments];

                        _wasSearchedSectionGroupsInSingleNotebook = true;  // нашли необходимые группы секций. 
                    }
                }
                else
                {
                    FormLogger.LogError(string.Format(BibleCommon.Resources.Constants.ConfiguratorWrongNotebookSelected + "\n" + errorText, notebookName, ContainerType.Single));
                }
            }
            else
            {
                FormLogger.LogMessage(BibleCommon.Resources.Constants.ConfiguratorNotebookNotDefined);
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
            tbCommentsPageName.Enabled = !chkDefaultPageNameParameters.Checked;
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
            LongProcessLogger.AbortedByUsers = true;
        }

        private void MainForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            BibleCommon.Services.Logger.Done();
            _oneNoteApp = null;

            if (_notebookParametersForm != null)
                _notebookParametersForm.Dispose();

            if (_loadForm != null)
                _loadForm.Dispose();
        }

        private void btnRelinkComments_Click(object sender, EventArgs e)
        {
            using (var manager = new RelinkAllBibleCommentsManager(_oneNoteApp, this))
            {
                manager.RelinkAllBibleComments();
            }
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
            btnClose.Enabled = false;
            btnApply.Enabled = false;
            System.Windows.Forms.Application.DoEvents();
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
            FormExtensions.SetControlPropertyThreadSafe(btnClose, "Enabled", true);
            FormExtensions.SetControlPropertyThreadSafe(btnApply, "Enabled", true);

            System.Windows.Forms.Application.DoEvents();
        }

        public void PerformProgressStep(string infoText)
        {
            FormExtensions.SetControlPropertyThreadSafe(lblProgressInfo, "Text", infoText);
            
            if (pbMain.Value < pbMain.Maximum)
                FormExtensions.SetControlPropertyThreadSafe(pbMain, "Value", pbMain.Value + 1);                
            
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
                {
                    using (var manager = new DeleteNotesPagesManager(_oneNoteApp, this))
                    {
                        manager.DeleteNotesPages();
                    }
                }
            }
        }

        private void btnResizeBibleTables_Click(object sender, EventArgs e)
        {
            using (SetWidthForm form = new SetWidthForm())
            {
                if (form.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    using (ResizeBibleManager manager = new ResizeBibleManager(_oneNoteApp, this))
                    {
                        manager.ResizeBiblePages(form.BiblePagesWidth);
                    }
                }
            }
        }

        private void btnBackup_Click(object sender, EventArgs e)
        {
            saveFileDialog.DefaultExt = ".zip";
            saveFileDialog.FileName = string.Format("{0}_backup_{1}", Constants.ToolsName, DateTime.Now.ToString("yyyy.MM.dd"));

            if (saveFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                BackupManager manager = new BackupManager(_oneNoteApp, this);

                manager.Backup(saveFileDialog.FileName);
            }
        }        

        private void tabPage1_Enter(object sender, EventArgs e)
        {
            if (!SettingsManager.Instance.CurrentModuleIsCorrect())            
                tbcMain.SelectedTab = tbcMain.TabPages[tabPage4.Name];
        }

        private void btnUploadModule_Click(object sender, EventArgs e)
        {
            try
            {
                btnUploadModule.Enabled = false;
                if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    bool moduleWasAdded;
                    bool needToReload = AddNewModule(openFileDialog.FileName, out moduleWasAdded);
                    if (needToReload)
                        ReLoadParameters(true);
                }
            }
            finally
            {
                btnUploadModule.Enabled = true;
            }
        }

        private void ReLoadModulesInfo()
        {
            pnModules.Controls.Clear();
            LoadModulesInfo();
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="needToLoadParameters"></param>
        /// <returns>true если новый модуль стал основным</returns>
        public bool AddNewModule(string filePath, out bool moduleWasAdded)
        {
            var preModuleInfo = ModulesManager.ReadModuleInfo(filePath);
            string moduleName = preModuleInfo.ShortName;            
            
            string destFilePath = Path.Combine(ModulesManager.GetModulesPackagesDirectory(), moduleName + Constants.FileExtensionIsbt);

            moduleWasAdded = true;
            bool canContinue = true;
            if (File.Exists(destFilePath))
            {
                var needToAsk = false;

                ModuleInfo existingModule = null;

                try
                {
                    existingModule = ModulesManager.GetModuleInfo(moduleName);
                    if (existingModule.Version.CompareTo(preModuleInfo.Version) > 0) 
                        needToAsk = true;
                }
                catch (InvalidModuleException)
                { }                   

                
                if (needToAsk 
                    && existingModule != null 
                    && MessageBox.Show(string.Format(BibleCommon.Resources.Constants.ModuleWithSameNameAlreadyExists, existingModule.Version, preModuleInfo.Version),
                                                BibleCommon.Resources.Constants.Warning, MessageBoxButtons.YesNo, 
                                                MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1) == System.Windows.Forms.DialogResult.No)
                {
                    canContinue = false;
                    moduleWasAdded = false;
                }
            }

            if (canContinue)
            {
                ModuleInfo module = null;

                try
                {
                    bool currentModuleIsCorrect = SettingsManager.Instance.CurrentModuleIsCorrect();  // а то может быть, что мы загрузили модуль, и он стал корретным, но UI не обновилось

                    module = ModulesManager.UploadModule(filePath, destFilePath, moduleName);

                    bool needToReload = false;

                    if (!currentModuleIsCorrect && module.Type == ModuleType.Bible)
                    {
                        SettingsManager.Instance.ModuleName = module.ShortName;
                        needToReload = true;
                    }
                    
                    ReLoadModulesInfo();

                    FormLogger.LogMessage(BibleCommon.Resources.Constants.ModuleSuccessfullyUploaded);

                    return needToReload;                    
                }
                catch (InvalidModuleException ex)
                {
                    FormLogger.LogError(ex);                    
                    Thread.Sleep(500);

                    if (module != null)
                        ModulesManager.DeleteModule(module.ShortName);
                    else
                        ModulesManager.DeleteModule(moduleName);

                    moduleWasAdded = false;
                }
            }

            return false;
        }
        

        private bool _wasLoadedModulesInfo = false;

        private void tabPage4_Enter(object sender, EventArgs e)
        {
            if (!_wasLoadedModulesInfo)
            {
                LoadModulesInfo();
            }

            btnSupplementalBibleManagement.Text = BibleCommon.Resources.Constants.SupplementalBibleManagement;
            btnDictionariesManagement.Text = BibleCommon.Resources.Constants.DictionariesManagement;
        }

        private static int GetModuleTypeWeight(ModuleType type)
        {
            switch (type)
            {
                case ModuleType.Bible:
                    return 0;
                default:
                    return 1;
            }
        }


        private bool _lblModulesBibleTitleWasAdded = false;
        private bool _lblModulesDictionariesTitleWasAdded = false;


        private const int MaxPnModulesHeight = 265;
        private void LoadModulesInfo()
        {            

            int top = 10;
            _lblModulesBibleTitleWasAdded = false;
            _lblModulesDictionariesTitleWasAdded = false;
            var modules = ModulesManager.GetModules();
            foreach (var module in modules.OrderBy(m => GetModuleTypeWeight(m.Type)).ThenBy(m => m.Name))
            {
                try
                {
                    ModulesManager.CheckModule(module);

                    top = SetModulesGroupTitle(module, top);

                    LoadModuleToUI(module, top);                    
                }
                catch (Exception ex)
                {
                    string moduleDirectory = ModulesManager.GetModuleDirectory(module.ShortName);
                    FormLogger.LogMessage(string.Format(BibleCommon.Resources.Constants.ModuleUploadError, moduleDirectory, ex.Message));
                    if (DeleteModuleWithConfirm(module.ShortName))
                        return;
                }

                top += 30;
            }

            if (top > MaxPnModulesHeight)
                top = MaxPnModulesHeight;

            pnModules.Height = top;
            btnUploadModule.Top = top + 50;

            if (modules.Where(m => m.Type == ModuleType.Bible).Count() > 0)
            {                
                btnUploadModule.Left = 31 + pnModules.Left;

                btnSupplementalBibleManagement.Top = btnUploadModule.Top;
                btnSupplementalBibleManagement.Left = btnUploadModule.Right + 6;
                btnSupplementalBibleManagement.Visible = true;

                btnDictionariesManagement.Top = btnUploadModule.Top;
                btnDictionariesManagement.Left = btnSupplementalBibleManagement.Right + 6;
                btnDictionariesManagement.Visible = true;

                lblMustUploadModule.Visible = false;
                lblMustSelectModule.Visible = !SettingsManager.Instance.CurrentModuleIsCorrect();
            }
            else
            {
                if (modules.Count == 0)
                    btnUploadModule.Top = 125;

                lblMustUploadModule.Top = btnUploadModule.Top - 20;                

                btnUploadModule.Left = (this.Width - btnUploadModule.Width) / 2;                

                lblMustUploadModule.Visible = true;
                lblMustSelectModule.Visible = false;
                btnSupplementalBibleManagement.Visible = false;
                btnDictionariesManagement.Visible = false;
            }
            
            _wasLoadedModulesInfo = true;
        }

        private int SetModulesGroupTitle(ModuleInfo module, int top)
        {
            Label lblTitle = null;
            if (module.Type == ModuleType.Bible)
            {
                if (!_lblModulesBibleTitleWasAdded)
                {
                    lblTitle = new Label() { Text = BibleCommon.Resources.Constants.BaseModules, Top = top + 10, Width = 600 };
                    _lblModulesBibleTitleWasAdded = true;
                }
            }
            else
            {
                if (!_lblModulesDictionariesTitleWasAdded)
                {
                    lblTitle = new Label() { Text = BibleCommon.Resources.Constants.AdditionalModules, Top = top + 10, Width = 600 };
                    _lblModulesDictionariesTitleWasAdded = true;
                }
            }

            if (lblTitle != null)
            {
                lblTitle.Font = new Font(lblTitle.Font, FontStyle.Bold);
                pnModules.Controls.Add(lblTitle);
                top += 35;
            }

            return top;
        }

        private void LoadModuleToUI(ModuleInfo moduleInfo, int top)
        {   
            var maximumModuleNameLength = 45;
            Label lblName = new Label();
            if (moduleInfo.Name.Length > maximumModuleNameLength)
                lblName.Text = moduleInfo.Name.Substring(0, maximumModuleNameLength) + "...";
            else
                lblName.Text = moduleInfo.Name;
            lblName.Top = top + 5;
            lblName.Left = 15;
            lblName.Width = 310;
            FormExtensions.SetToolTip(lblName, BibleCommon.Resources.Constants.ModuleDisplayName);
            pnModules.Controls.Add(lblName);

            Label lblVersion = new Label();
            lblVersion.Text = moduleInfo.Version;
            lblVersion.Top = top + 5;
            lblVersion.Left = 330;
            lblVersion.Width = 25;
            FormExtensions.SetToolTip(lblVersion, BibleCommon.Resources.Constants.ModuleVersion);
            pnModules.Controls.Add(lblVersion);

            if (moduleInfo.Type == ModuleType.Bible)
            {
                CheckBox cbIsActive = new CheckBox();
                cbIsActive.AutoCheck = false;
                cbIsActive.Checked = SettingsManager.Instance.ModuleName == moduleInfo.ShortName;
                cbIsActive.Top = top;
                cbIsActive.Left = 355;
                cbIsActive.Width = 20;
                FormExtensions.SetToolTip(cbIsActive, BibleCommon.Resources.Constants.ModuleIsActive);
                pnModules.Controls.Add(cbIsActive);
            }

            
            Button btnInfo = new Button();
            btnInfo.Text = "?";
            btnInfo.Tag = moduleInfo;
            btnInfo.Top = top;
            btnInfo.Left = 375;
            btnInfo.Width = 20;
            btnInfo.Click += new EventHandler(btnModuleInfo_Click);
            FormExtensions.SetToolTip(btnInfo, BibleCommon.Resources.Constants.ModuleInformation);
            pnModules.Controls.Add(btnInfo);
            

            if (moduleInfo.Type == ModuleType.Bible)
            {
                Button btnUseThisModule = new Button();
                btnUseThisModule.Text = GetBtnModuleManagementText(moduleInfo.Type);
                btnUseThisModule.Enabled = moduleInfo.Type == ModuleType.Bible ? SettingsManager.Instance.ModuleName != moduleInfo.ShortName : true;
                btnUseThisModule.Tag = moduleInfo;
                btnUseThisModule.Top = top;
                btnUseThisModule.Left = 400;
                btnUseThisModule.Width = 185;
                btnUseThisModule.Click += new EventHandler(btnUseThisModule_Click);
                pnModules.Controls.Add(btnUseThisModule);
            }

            Button btnDel = new Button();
            btnDel.Image = BibleConfigurator.Properties.Resources.del;
            btnDel.Enabled = SettingsManager.Instance.ModuleName != moduleInfo.ShortName;            
            btnDel.Tag = moduleInfo.ShortName;
            btnDel.Top = top;            
            btnDel.Left = 590;
            btnDel.Width = btnDel.Height;
            btnDel.Click += new EventHandler(btnDeleteModule_Click);
            FormExtensions.SetToolTip(btnDel, BibleCommon.Resources.Constants.DeleteThisModule);
            pnModules.Controls.Add(btnDel);            
        }       

        private string GetBtnModuleManagementText(ModuleType moduleType)
        {
            switch (moduleType)
            {
                case ModuleType.Bible: 
                    return BibleCommon.Resources.Constants.UseThisModule; 
                case ModuleType.Strong:
                    return BibleCommon.Resources.Constants.SupplementalBibleManagement; 
                case ModuleType.Dictionary:
                    return BibleCommon.Resources.Constants.DictionariesManagement;
                default: 
                    throw new NotSupportedException(moduleType.ToString());
            }
        }

        void btnModuleInfo_Click(object sender, EventArgs e)
        {
            var btn = (Button)sender;
            var moduleInfo = (ModuleInfo)btn.Tag;

            if (moduleInfo.Type == ModuleType.Dictionary)
            {
                MessageBox.Show(!string.IsNullOrEmpty(moduleInfo.Description) ? moduleInfo.Description : moduleInfo.Name, 
                    BibleCommon.Resources.Constants.ModuleInformation, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                using (AboutModuleForm f = new AboutModuleForm(moduleInfo.ShortName, false))
                {
                    f.ShowDialog();
                }
            }
        }

        void btnUseThisModule_Click(object sender, EventArgs e)
        {
            var btn = (Button)sender;
            var moduleInfo = (ModuleInfo)btn.Tag;

            switch (moduleInfo.Type)
            {
                case ModuleType.Bible:
                    bool canContinue = true;

                    if (!string.IsNullOrEmpty(SettingsManager.Instance.NotebookId_Bible) && OneNoteUtils.NotebookExists(_oneNoteApp, SettingsManager.Instance.NotebookId_Bible, true)
                        && SettingsManager.Instance.CurrentModuleIsCorrect())
                    {
                        if (MessageBox.Show(BibleCommon.Resources.Constants.ChangeModuleWarning, BibleCommon.Resources.Constants.Warning,
                            MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == System.Windows.Forms.DialogResult.No)
                            canContinue = false;
                    }

                    if (canContinue)
                    {                        
                        SettingsManager.Instance.ModuleName = moduleInfo.ShortName;
                        ReLoadModulesInfo();
                        ReLoadParameters(true);
                    }
                    break;
                case ModuleType.Strong:
                    ShowSupplementalBibleManagementForm();
                    break;
                case ModuleType.Dictionary:
                    ShowDictionariesManagementForm();
                    break;
            }
        }        

        private void ReLoadParameters(bool needToSaveSettings)
        {
            _loadForm.SetDesktopLocation(this.Left - 5, this.Top - 5);
            _loadForm.Show();
            try
            {
                LoadParameters(ModulesManager.GetCurrentModuleInfo(), needToSaveSettings);
            }
            finally
            {
                _loadForm.Hide();
            }
        }

        void btnDeleteModule_Click(object sender, EventArgs e)
        {
            var btn = (Button)sender;
            var moduleName = (string)btn.Tag;

            if (SettingsManager.Instance.SupplementalBibleModules.Contains(moduleName))
                FormLogger.LogError(BibleCommon.Resources.Constants.ModuleCannotBeDeleted_SupplementalBibleModule);
            else if (SettingsManager.Instance.DictionariesModules.Any(m => m.ModuleName == moduleName))
                FormLogger.LogError(BibleCommon.Resources.Constants.ModuleCannotBeDeleted_DictionaryModule);
            else
                DeleteModuleWithConfirm(moduleName);
        }

        private bool DeleteModuleWithConfirm(string moduleName)
        {
            if (MessageBox.Show(BibleCommon.Resources.Constants.DeleteThisModuleQuestion, BibleCommon.Resources.Constants.Warning, MessageBoxButtons.YesNo, MessageBoxIcon.Warning)  
               == System.Windows.Forms.DialogResult.Yes)
            {
                ModulesManager.DeleteModule(moduleName);

                ReLoadModulesInfo();
                return true;
            }

            return false;
        }

        private void hlModules_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start(BibleCommon.Resources.Constants.WebSiteUrl + "/modules.htm");
        }

        private void btnSupplementalBibleManagement_Click(object sender, EventArgs e)
        {
            ShowSupplementalBibleManagementForm();
        }

        private void ShowSupplementalBibleManagementForm()
        {
            using (var form = new SupplementalBibleForm(_oneNoteApp, this))
            {
                form.ShowDialog();
            }
        }

        private void btnDictionariesManagement_Click(object sender, EventArgs e)
        {
            ShowDictionariesManagementForm();
        }

        private void ShowDictionariesManagementForm()
        {
            using (var form = new DictionaryModulesForm(_oneNoteApp, this))
            {
                form.ShowDialog();
            }
        }                          
    }
}
