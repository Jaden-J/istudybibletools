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

namespace BibleConfigurator
{
    public abstract partial class BaseSupplementalForm : Form
    {
        protected Microsoft.Office.Interop.OneNote.Application OneNoteApp { get; set; }
        protected MainForm MainForm { get; set; }
        protected List<ModuleInfo> Modules { get; set; }
        protected Button BtnAddNewModule { get; set; }
        protected int TopControlsPosition { get; set; }
        protected ComboBox CbModule { get; set; }
        protected CustomFormLogger Logger { get; set; }
        protected bool NeedToCommitChanges { get; set; }
        protected bool WasLoaded { get; set; }
        protected bool InProgress { get; set; }

        protected abstract string GetValidSupplementalNotebookId();
        protected abstract void ClearSupplementalModulesInSettingsStorage();
        protected abstract int GetSupplementalModulesCount();
        protected abstract bool SupplementalModuleAlreadyAdded(string moduleShortName);
        protected abstract string FormDescription { get; }
        protected abstract List<string> CommitChanges(ModuleInfo selectedModuleInfo);
        protected abstract string GetSupplementalModuleName(int index);
        protected abstract bool CanModuleBeDeleted(int index);
        protected abstract void DeleteModule(string moduleShortName);
        protected abstract string CloseSupplementalNotebookConfirmText { get; }
        protected abstract void CloseSupplementalNotebook();
        protected abstract bool IsModuleSupported(ModuleInfo moduleInfo);        

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
            Modules = ModulesManager.GetModules();
            TopControlsPosition = 10;
            Logger = new CustomFormLogger(MainForm);

            InitializeComponent();            
        }

        private void SupplementalBibleForm_Load(object sender, EventArgs e)
        {
            if (!SettingsManager.Instance.IsConfigured(OneNoteApp))
            {
                FormLogger.LogError(BibleCommon.Resources.Constants.Error_SystemIsNotConfigures);
                Close();
            }

            if (!BibleParallelTranslationManager.IsModuleSupported(SettingsManager.Instance.CurrentModule))
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
        }

        private void LoadFormData()
        {
            WasLoaded = false;  

            chkUseSupplementalBible.Checked = !string.IsNullOrEmpty(GetValidSupplementalNotebookId());
            if (!chkUseSupplementalBible.Checked)
                ClearSupplementalModulesInSettingsStorage(); // на всякий пожарный

            chkUseSupplementalBible_CheckedChanged(this, null);

            WasLoaded = true;
        }

        private void LoadUI()
        {            
            pnModules.Controls.Clear();
            TopControlsPosition = 10;
            NeedToCommitChanges = false;

            if (!chkUseSupplementalBible.Checked && GetSupplementalModulesCount() == 0)
            {
                Label lblDescription = new Label();
                lblDescription.Text = FormDescription;
                lblDescription.Top = TopControlsPosition;
                lblDescription.Width = 260;
                lblDescription.Height = 100;
                lblDescription.Left = 20;
                pnModules.Controls.Add(lblDescription);
            }
            else
            {
                GenerateNewModuleButton();

                if (GetSupplementalModulesCount() == 0)
                    _btnAddNewModule_Click(this, null);
                else
                    LoadModules();
            }
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
                BtnAddNewModule.Text = BibleCommon.Resources.Constants.Apply;
                BtnAddNewModule.Image = Properties.Resources.apply;
                BtnAddNewModule.ImageAlign = ContentAlignment.MiddleRight;
                BtnAddNewModule.Width = 85;
                NeedToCommitChanges = true;
            }
            else
            {
                CommitChanges();               
            }
        }

        private void CommitChanges()
        {
            EnableUI(false);
            InProgress = true;

            BibleCommon.Services.Logger.LogMessage("Start work with supplemental modules");
            var dt = DateTime.Now;

            try
            {
                var selectedModuleInfo = ((ModuleInfo)CbModule.SelectedItem);

                var errors = CommitChanges(selectedModuleInfo);               

                BibleCommon.Services.Logger.LogMessage("Finish work with supplemental modules. Elapsed time = '{0}'", DateTime.Now - dt);

                if (errors.Count > 0)
                {
                    var errorsForm = new BibleCommon.UI.Forms.ErrorsForm(errors);
                    errorsForm.ShowDialog();
                }
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
        }

        private void EnableUI(bool enabled)
        {
            pnModules.Enabled = enabled;
            chkUseSupplementalBible.Enabled = enabled;
            btnOk.Enabled = enabled;
            btnCancel.Enabled = enabled;
            btnSBFolder.Enabled = enabled;
        }

        private void LoadModules()
        {
            for (int i = 0; i < GetSupplementalModulesCount(); i++)
            {
                AddModuleRow(Modules.First(m => m.ShortName == GetSupplementalModuleName(i)), i, TopControlsPosition);
                TopControlsPosition += 30;
            }

            BtnAddNewModule.Top = TopControlsPosition;
        }

        private void AddModuleRow(ModuleInfo moduleInfo, int index, int top)
        {
            Label lblName = new Label();
            lblName.Text = moduleInfo.Name;
            lblName.Top = top + 5;
            lblName.Left = 0;
            lblName.Width = 245;
            pnModules.Controls.Add(lblName);

            Button btnDel = new Button();
            btnDel.Image = BibleConfigurator.Properties.Resources.del;
            btnDel.Enabled = CanModuleBeDeleted(index);
            FormExtensions.SetToolTip(btnDel, BibleCommon.Resources.Constants.DeleteThisModule);
            btnDel.Tag = moduleInfo.ShortName;
            btnDel.Top = top;
            btnDel.Left = 248;
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
            InProgress = true;

            var result = false;

            if (MessageBox.Show(BibleCommon.Resources.Constants.DeleteThisModuleQuestion, BibleCommon.Resources.Constants.Warning, MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
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
                    MainForm.ExternalProcessingDone(BibleCommon.Resources.Constants.ErrorOccurred);
                }

                LoadFormData();
                result = true;
            }

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
                btnSBFolder.Visible = false;

                if (!string.IsNullOrEmpty(sbNotebookId))
                {
                    if (MessageBox.Show(CloseSupplementalNotebookConfirmText,
                        BibleCommon.Resources.Constants.Warning, MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
                        == System.Windows.Forms.DialogResult.Yes)
                    {
                        CloseSupplementalNotebook();
                    }
                    else
                    {
                        chkUseSupplementalBible.Checked = !chkUseSupplementalBible.Checked;
                        needToUpdate = false;
                    }
                }
            }
            else if (chkUseSupplementalBible.Checked && string.IsNullOrEmpty(sbNotebookId))
            {
                btnSBFolder.Visible = true;
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

        private void AddModulesComboBox()
        {            
            CbModule = new ComboBox();
            CbModule.DropDownStyle = ComboBoxStyle.DropDownList;
            CbModule.Width = 245;
            CbModule.Top = TopControlsPosition;
            CbModule.ValueMember = "Name";

            TopControlsPosition = TopControlsPosition + 30;

            foreach (var moduleInfo in Modules)
            {
                if (IsModuleSupported(moduleInfo) && !SupplementalModuleAlreadyAdded(moduleInfo.ShortName))
                    CbModule.Items.Add(moduleInfo);
            }

            if (CbModule.Items.Count > 0)
                CbModule.SelectedIndex = 0;

            pnModules.Controls.Add(CbModule);
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            if (NeedToCommitChanges)
                CommitChanges();

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

    }
}
