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
using BibleCommon.Common;
using Microsoft.Office.Interop.OneNote;
using System.Xml;
using System.Xml.XPath;
using System.Xml.Linq;

namespace BibleConfigurator
{
    public abstract partial class BaseSupplementalForm : Form
    {
        protected Microsoft.Office.Interop.OneNote.Application OneNoteApp { get; set; }
        protected MainForm MainForm { get; set; }
        protected List<ModuleInfo> Modules { get; set; }        
        protected Button BtnAddNewModule { get; set; }
        protected int TopControlsPosition { get; set; }
        protected int LeftControlsPosition { get; set; }
        protected ComboBox CbModule { get; set; }
        protected CustomFormLogger Logger { get; set; }
        protected bool NeedToCommitChanges { get; set; }
        protected bool WasLoaded { get; set; }
        protected bool InProgress { get; set; }
        protected Dictionary<string, string> ExistingNotebooks { get; set; }

        private Label LblDescription { get; set; }

        protected abstract string GetFormText();
        protected abstract string GetChkUseText();
        protected abstract string GetValidSupplementalNotebookId();        
        protected abstract int GetSupplementalModulesCount();
        protected abstract void ClearSupplementalModules();
        protected abstract bool SupplementalModuleAlreadyAdded(string moduleShortName);
        protected abstract string FormDescription { get; }
        protected abstract List<string> CommitChanges(ModuleInfo selectedModuleInfo);        
        protected abstract string GetSupplementalModuleName(int index);
        protected abstract bool CanModuleBeDeleted(ModuleInfo moduleInfo, int index);
        protected abstract bool CanModuleBeAdded(ModuleInfo moduleInfo);
        protected abstract void DeleteModule(string moduleShortName);
        protected abstract string CloseSupplementalNotebookQuestionText { get; }
        protected abstract void CloseSupplementalNotebook();
        protected abstract bool IsModuleSupported(ModuleInfo moduleInfo);
        protected abstract bool IsBaseModuleSupported();
        protected abstract string DeleteModuleQuestionText { get; }
        protected abstract bool CanNotebookBeClosed();
        protected abstract string NotebookCannotBeClosedText { get; }
        protected abstract string EmbeddedModulesKey { get; }
        protected abstract string NotebookIsNotSupplementalBibleMessage { get; }
        protected abstract string SupplementalNotebookWasAddedMessage { get; }
        protected abstract void SaveSupplementalNotebookSettings(string notebookId);
        protected abstract List<string> SaveEmbeddedModuleSettings(EmbeddedModuleInfo embeddedModuleInfo, ModuleInfo moduleInfo, XElement pageEl);


        protected FolderBrowserDialog FolderBrowserDialog
        {
            get
            {
                return folderBrowserDialog;
            }
        }

        public BaseSupplementalForm(Microsoft.Office.Interop.OneNote.Application oneNoteApp, MainForm form)
        {
            OneNoteApp = oneNoteApp;
            MainForm = form;
            Modules = ModulesManager.GetModules(true);
            TopControlsPosition = 10;
            Logger = new CustomFormLogger(MainForm);

            this.SetFormUICulture();

            InitializeComponent();            
        }

        private void SupplementalBibleForm_Load(object sender, EventArgs e)
        {
            try
            {
                if (!SettingsManager.Instance.IsConfigured(OneNoteApp))
                {
                    FormLogger.LogError(BibleCommon.Resources.Constants.Error_SystemIsNotConfigures);
                    Close();
                }

                if (!IsBaseModuleSupported())
                {
                    FormLogger.LogError(string.Format(BibleCommon.Resources.Constants.BaseModuleIsNotSupported,
                                            SettingsManager.Instance.CurrentModule.Version, BibleParallelTranslationManager.SupportedModuleMinVersion));
                    Close();
                }

                LoadFormData();

                string defaultNotebookFolderPath;
                OneNoteApp.GetSpecialLocation(SpecialLocation.slDefaultNotebookFolder, out defaultNotebookFolderPath);
                folderBrowserDialog.SelectedPath = defaultNotebookFolderPath;
                folderBrowserDialog.Description = BibleCommon.Resources.Constants.ConfiguratorSetNotebookFolder;
                folderBrowserDialog.ShowNewFolderButton = true;

                FormExtensions.SetToolTip(btnSBFolder, BibleCommon.Resources.Constants.DefineNotebookDirectory);

                this.Text = GetFormText();
                chkUseSupplementalBible.Text = GetChkUseText();
            }
            catch (Exception ex)
            {                
                FormLogger.LogError(ex);
            }
        }

        private void LoadFormData()
        {
            WasLoaded = false;  

            bool supplementalNotebookIsInUse = !string.IsNullOrEmpty(GetValidSupplementalNotebookId());            

            if (chkUseSupplementalBible.Checked != supplementalNotebookIsInUse)
                chkUseSupplementalBible.Checked = supplementalNotebookIsInUse;
            else
                chkUseSupplementalBible_CheckedChanged(this, null);

            WasLoaded = true;
        }

        private void LoadUI()
        {
            ResetControls();    
            
            NeedToCommitChanges = false;

            if (!chkUseSupplementalBible.Checked && GetSupplementalModulesCount() == 0)
            {
                LblDescription = new Label();
                LblDescription.Text = FormDescription;
                LblDescription.Top = TopControlsPosition;
                LblDescription.Width = 360;
                LblDescription.Height = 150;
                LblDescription.Left = LeftControlsPosition;                
                pnModules.Controls.Add(LblDescription);
                pnModules.Enabled = true;
                btnSBFolder.Visible = false;
                rbCreateNew.Visible = false;
                rbUseExisting.Visible = false;
                cbExistingNotebooks.Visible = false;
            }
            else
            {
                if (AreThereModulesToAdd())
                    GenerateNewModuleButton();

                var notebookIsInUse = GetSupplementalModulesCount() != 0 && !string.IsNullOrEmpty(GetValidSupplementalNotebookId());

                if (!notebookIsInUse)
                {
                    TopControlsPosition += 20;
                    LeftControlsPosition += 24;
                    _btnAddNewModule_Click(this, null);
                    LoadExistingNotebooks();
                }
                else
                {                    
                    LoadModules();
                }

                btnSBFolder.Visible = !notebookIsInUse;
                rbCreateNew.Visible = !notebookIsInUse;
                rbUseExisting.Visible = !notebookIsInUse;
                cbExistingNotebooks.Visible = !notebookIsInUse;
            }
        }

        private void LoadExistingNotebooks()
        {
            if (ExistingNotebooks == null)
                ExistingNotebooks = OneNoteUtils.GetExistingNotebooks(OneNoteApp);

            cbExistingNotebooks.DataSource = ExistingNotebooks.Values.ToList();            
        }

        private void ResetControls()
        {
            pnModules.Controls.Clear();
            pnModules.Controls.Add(rbCreateNew);
            pnModules.Controls.Add(rbUseExisting);
            pnModules.Controls.Add(cbExistingNotebooks);

            TopControlsPosition = 10;
            LeftControlsPosition = 0;
            rbCreateNew.Checked = true;
            cbExistingNotebooks.Enabled = false;
        }

        private void GenerateNewModuleButton()
        {
            BtnAddNewModule = new Button();
            BtnAddNewModule.Image = BibleConfigurator.Properties.Resources.plus;
            FormExtensions.SetToolTip(BtnAddNewModule, BibleCommon.Resources.Constants.AddSupplementalModule);
            BtnAddNewModule.Click += new EventHandler(_btnAddNewModule_Click);
            BtnAddNewModule.Width = BtnAddNewModule.Height;            
            BtnAddNewModule.Enabled = GetSupplementalModulesCount() < Modules.Count;            
            pnModules.Controls.Add(BtnAddNewModule);
        }

        void _btnAddNewModule_Click(object sender, EventArgs e)
        {
            if (!NeedToCommitChanges)
            {
                AddModulesComboBox();

                BtnAddNewModule.TextAlign = ContentAlignment.MiddleLeft;
                BtnAddNewModule.Top = TopControlsPosition;
                BtnAddNewModule.Left = LeftControlsPosition;
                BtnAddNewModule.Text = BibleCommon.Resources.Constants.Apply;
                BtnAddNewModule.Image = Properties.Resources.apply;
                BtnAddNewModule.ImageAlign = ContentAlignment.MiddleRight;
                BtnAddNewModule.Width = 85;
                NeedToCommitChanges = true;
            }
            else
            {
                CommitChanges(false);               
            }
        }

        private void CommitChanges(bool closeAfter)
        {
            bool useExistingNotebook = rbUseExisting.Enabled && rbUseExisting.Checked;
            EnableUI(false);
            InProgress = true;

            BibleCommon.Services.Logger.LogMessage("Start work with supplemental modules");
            var dt = DateTime.Now;

            bool doNotClose = false;

            try
            {
                List<string> errors;
                if (!useExistingNotebook)
                {
                    var selectedModuleInfo = ((ModuleInfo)CbModule.SelectedItem);
                    errors = CommitChanges(selectedModuleInfo);
                }
                else
                {
                    errors = TryToUseExistingNotebook(ExistingNotebooks.First(n => n.Value == (string)cbExistingNotebooks.SelectedValue).Key, (string)cbExistingNotebooks.SelectedValue);
                }

                BibleCommon.Services.Logger.LogMessage("Finish work with supplemental modules. Elapsed time = '{0}'", DateTime.Now - dt);

                if (errors.Count > 0)
                {
                    using (var errorsForm = new BibleCommon.UI.Forms.ErrorsForm(errors))
                    {
                        errorsForm.ShowDialog();
                    }
                }
            }
            catch (InvalidNotebookException ex)
            {
                FormLogger.LogError(ex);
                MainForm.ExternalProcessingDone(ex.Message);
                doNotClose = true;
            }
            catch (ProcessAbortedByUserException)
            {
                BibleCommon.Services.Logger.LogMessage("Process aborted by user");
                MainForm.ExternalProcessingDone(BibleCommon.Resources.Constants.ProcessAbortedByUser);
            }
            catch (Exception ex)
            {
                BibleCommon.Services.Logger.LogError(ex.ToString());
                MainForm.ExternalProcessingDone(string.Format("{0}: {1}", BibleCommon.Resources.Constants.ErrorOccurred, ex.Message));
            }                        

            EnableUI(true);
            InProgress = false;

            LoadUI();

            if (closeAfter && !doNotClose)
                this.Close();
        }

        private void EnableUI(bool enabled)
        {
            pnModules.Enabled = enabled;            
            chkUseSupplementalBible.Enabled = enabled;
            btnOk.Enabled = enabled;            
            btnSBFolder.Enabled = enabled;            
        }

        private void LoadModules()
        {
            for (int i = 0; i < GetSupplementalModulesCount(); i++)
            {
                var module = Modules.First(m => m.ShortName == GetSupplementalModuleName(i));
                if (IsModuleSupported(module))
                {
                    AddModuleRow(module, i, TopControlsPosition);
                    TopControlsPosition += 30;
                }
            }

            if (BtnAddNewModule != null)
                BtnAddNewModule.Top = TopControlsPosition;
        }

        private void AddModuleRow(ModuleInfo moduleInfo, int index, int top)
        {
            Label lblName = new Label();
            lblName.Text = moduleInfo.DisplayName;
            lblName.Top = top + 5;
            lblName.Left = 0;
            lblName.Width = 345;
            pnModules.Controls.Add(lblName);

            Button btnDel = new Button();
            btnDel.Image = BibleConfigurator.Properties.Resources.del;
            btnDel.Enabled = CanModuleBeDeleted(moduleInfo, index);
            FormExtensions.SetToolTip(btnDel, BibleCommon.Resources.Constants.DeleteThisModule);
            btnDel.Tag = moduleInfo.ShortName;
            btnDel.Top = top;
            btnDel.Left = 348;
            btnDel.Width = btnDel.Height;
            btnDel.Click += new EventHandler(btnDeleteModule_Click);
            pnModules.Controls.Add(btnDel);
        }

        void btnDeleteModule_Click(object sender, EventArgs e)
        {
            var btn = (Button)sender;
            var moduleName = (string)btn.Tag;

            DeleteModuleWithConfirm(moduleName);
        }

        private bool DeleteModuleWithConfirm(string moduleName)
        {
            EnableUI(false);
            InProgress = true;

            var result = false;

            if (MessageBox.Show(DeleteModuleQuestionText, BibleCommon.Resources.Constants.Warning, MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
                == System.Windows.Forms.DialogResult.Yes)
            {
                try
                {
                    DeleteModule(moduleName);  
                }
                catch (ProcessAbortedByUserException)
                {
                    BibleCommon.Services.Logger.LogMessage("Process aborted by user");
                    MainForm.ExternalProcessingDone(BibleCommon.Resources.Constants.ProcessAbortedByUser);
                }
                catch (Exception ex)
                {
                    BibleCommon.Services.Logger.LogError(ex.ToString());
                    MainForm.ExternalProcessingDone(string.Format("{0}: {1}", BibleCommon.Resources.Constants.ErrorOccurred, ex.Message));
                }

                LoadFormData();
                result = true;
            }

            EnableUI(true);
            InProgress = false;
            return result;
        }  


        private void SupplementalBibleForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            OneNoteApp = null;
            MainForm = null;
            Logger.Dispose();
        }

        private void chkUseSupplementalBible_CheckedChanged(object sender, EventArgs e)
        {
            bool needToUpdate = true;

            string sbNotebookId = GetValidSupplementalNotebookId();

            if (WasLoaded && !chkUseSupplementalBible.Checked)
            {
                if (!string.IsNullOrEmpty(sbNotebookId))
                {
                    if (CanNotebookBeClosed() && MessageBox.Show(CloseSupplementalNotebookQuestionText,
                        BibleCommon.Resources.Constants.Warning, MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
                        == System.Windows.Forms.DialogResult.Yes)
                    {
                        CloseSupplementalNotebook();
                    }
                    else
                    {
                        chkUseSupplementalBible.Checked = !chkUseSupplementalBible.Checked;
                        needToUpdate = false;

                        if (!CanNotebookBeClosed())
                            FormLogger.LogError(NotebookCannotBeClosedText);
                    }
                }
            }
            else if (chkUseSupplementalBible.Checked && string.IsNullOrEmpty(sbNotebookId))
            {
                if (!Modules.Any(m => IsModuleSupported(m)))
                {
                    FormLogger.LogError(BibleCommon.Resources.Constants.SupportedModulesNotFound);

                    chkUseSupplementalBible.Checked = !chkUseSupplementalBible.Checked;
                    needToUpdate = false;
                }                
            }
            
            if (needToUpdate)
            {
                pnModules.Enabled = chkUseSupplementalBible.Checked;
                LoadUI();
            }
        }

        private void SupplementalBibleForm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                Close();
        }

        private bool AreThereModulesToAdd()
        {
            return Modules.Any(m => IsModuleSupported(m) && !SupplementalModuleAlreadyAdded(m.ShortName));            
        }

        private void AddModulesComboBox()
        {            
            CbModule = new ComboBox();
            CbModule.DropDownStyle = ComboBoxStyle.DropDownList;
            CbModule.Width = 345;
            CbModule.Top = TopControlsPosition;
            CbModule.Left = LeftControlsPosition;
            CbModule.ValueMember = "DisplayName";

            TopControlsPosition = TopControlsPosition + 30;

            foreach (var moduleInfo in Modules)
            {
                if (IsModuleSupported(moduleInfo) && !SupplementalModuleAlreadyAdded(moduleInfo.ShortName) && CanModuleBeAdded(moduleInfo))
                    CbModule.Items.Add(moduleInfo);
            }

            if (CbModule.Items.Count > 0)
                CbModule.SelectedIndex = 0;

            pnModules.Controls.Add(CbModule);            
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            if (NeedToCommitChanges)
                CommitChanges(true);
            else
                this.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void SupplementalBibleForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (InProgress)
            {
                if (MessageBox.Show(BibleCommon.Resources.Constants.AbortTheOperation, 
                        BibleCommon.Resources.Constants.Warning, MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == System.Windows.Forms.DialogResult.No)
                    e.Cancel = true;
                else
                    Logger.AbortedByUsers = true;
            }
        }

        private void btnSBFolder_Click(object sender, EventArgs e)
        {
            folderBrowserDialog.ShowDialog();            
        }

        private void rbUseExisting_CheckedChanged(object sender, EventArgs e)
        {
            if (!((RadioButton)sender).Checked)
            {
                FormExtensions.EnableAll(true, pnModules.Controls);
                cbExistingNotebooks.Enabled = false;
            }
            else
            {
                FormExtensions.EnableAll(false, pnModules.Controls, rbCreateNew, rbUseExisting, cbExistingNotebooks);
                cbExistingNotebooks.Enabled = true;
            }
        }


        protected List<string> TryToUseExistingNotebook(string notebookId, string notebookName)
        {
            XmlNamespaceManager xnm;

            var result = new List<string>();
            ClearSupplementalModules();            
                        
            var xDoc = OneNoteUtils.GetHierarchyElement(OneNoteApp, notebookId, HierarchyScope.hsPages, out xnm);
            var pagesDocs = xDoc.Root.XPathSelectElements("//one:Page", xnm);
            int pagesCount = pagesDocs.Count();

            Logger.Preffix = BibleCommon.Resources.Constants.ProcessPage + " ";
            MainForm.PrepareForExternalProcessing(pagesCount, 1, string.Empty);

            foreach (var pageEl in pagesDocs)
            {
                var pageId = (string)pageEl.Attribute("ID");
                var pageName = (string)pageEl.Attribute("name");

                Logger.LogMessage(pageName);

                var embeddedModulesInfo_string = OneNoteUtils.GetPageMetaData(OneNoteApp, pageEl, EmbeddedModulesKey, xnm);
                if (!string.IsNullOrEmpty(embeddedModulesInfo_string))
                {
                    var embeddedModulesInfo = EmbeddedModuleInfo.Deserialize(embeddedModulesInfo_string);

                    foreach (var embeddedModuleInfo in embeddedModulesInfo)
                    {
                        var module = Modules.FirstOrDefault(m => m.ShortName == embeddedModuleInfo.ModuleName);

                        if (module == null)
                            throw new InvalidNotebookException(string.Format(BibleCommon.Resources.Constants.ModuleIsNotInstalled, embeddedModuleInfo.ModuleName));

                        if (!SupplementalModuleAlreadyAdded(embeddedModuleInfo.ModuleName))
                        {
                            result.AddRange(SaveEmbeddedModuleSettings(embeddedModuleInfo, module, pageEl));
                        }
                    }
                }
            }

            Logger.Preffix = string.Empty;

            if (GetSupplementalModulesCount() == 0)            
                throw new InvalidNotebookException(string.Format(NotebookIsNotSupplementalBibleMessage, notebookName));            
            else
            {
                MainForm.ExternalProcessingDone(SupplementalNotebookWasAddedMessage);
                SaveSupplementalNotebookSettings(notebookId);                
            }

            return result;
        }
    }
}
