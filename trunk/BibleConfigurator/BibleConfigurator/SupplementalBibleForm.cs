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

namespace BibleConfigurator
{
    public partial class SupplementalBibleForm : Form
    {
        private Microsoft.Office.Interop.OneNote.Application _oneNoteApp;
        private MainForm _form;
        private List<ModuleInfo> _modules;
        private Button _btnAddNewModule;
        private int _top;
        private ComboBox _cbModule;
        private CustomFormLogger _logger;
        private bool _needToCommitChanges = false;
        private bool _wasLoaded = false;

        public SupplementalBibleForm(Microsoft.Office.Interop.OneNote.Application oneNoteApp, MainForm form)
        {
            _oneNoteApp = oneNoteApp;
            _form = form;
            _modules = ModulesManager.GetModules();
            _top = 10;
            _logger = new CustomFormLogger(_form);

            InitializeComponent();            
        }

        private void SupplementalBibleForm_Load(object sender, EventArgs e)
        {
            LoadFormData();
        }

        private void LoadFormData()
        {
            _wasLoaded = false;  

            chkUseSupplementalBible.Checked = !string.IsNullOrEmpty(SettingsManager.Instance.GetValidSupplementalBibleNotebookId(_oneNoteApp));

            chkUseSupplementalBible_CheckedChanged(this, null);

            _wasLoaded = true;
        }

        private void LoadUI()
        {
            pnModules.Controls.Clear();
            _top = 10;
            _needToCommitChanges = false;

            if (!chkUseSupplementalBible.Checked && SettingsManager.Instance.SupplementalBibleModules.Count == 0)
            {
                Label lblDescription = new Label();
                lblDescription.Text = 
@"Здесь Вы можете управлять справочной Библией. Обратите внимание, ";
                lblDescription.Top = _top;
                lblDescription.Width = 250;
                lblDescription.Height = 100;
                lblDescription.Left = 30;
                pnModules.Controls.Add(lblDescription);
            }
            else
            {
                GenerateNewModuleButton();

                if (SettingsManager.Instance.SupplementalBibleModules.Count == 0)
                    _btnAddNewModule_Click(this, null);
                else
                    LoadModules();
            }
        }

        private void GenerateNewModuleButton()
        {
            _btnAddNewModule = new Button();
            _btnAddNewModule.Image = BibleConfigurator.Properties.Resources.plus;
            FormExtensions.SetToolTip(_btnAddNewModule, BibleCommon.Resources.Constants.AddSupplementalModule);
            _btnAddNewModule.Click += new EventHandler(_btnAddNewModule_Click);
            _btnAddNewModule.Width = _btnAddNewModule.Height;            
            _btnAddNewModule.Enabled = SettingsManager.Instance.SupplementalBibleModules.Count < _modules.Count;            
            pnModules.Controls.Add(_btnAddNewModule);
        }

        void _btnAddNewModule_Click(object sender, EventArgs e)
        {
            if (!_needToCommitChanges)
            {
                AddModulesComboBox();

                _btnAddNewModule.TextAlign = ContentAlignment.MiddleLeft;
                _btnAddNewModule.Top = _top;
                _btnAddNewModule.Text = BibleCommon.Resources.Constants.Apply;
                _btnAddNewModule.Image = Properties.Resources.apply;
                _btnAddNewModule.ImageAlign = ContentAlignment.MiddleRight;
                _btnAddNewModule.Width = 85;
                _needToCommitChanges = true;
            }
            else
            {
                CommitChanges();               
            }
        }

        private void CommitChanges()
        {
            EnableUI(false);

            try
            {
                var selectedModuleInfo = ((ModuleInfo)_cbModule.SelectedItem);
                BibleParallelTranslationConnectionResult result;

                if (SettingsManager.Instance.SupplementalBibleModules.Count > 0)
                {
                    int chaptersCount = ModulesManager.GetBibleChaptersCount(SettingsManager.Instance.SupplementalBibleModules.First());
                    _form.PrepareForExternalProcessing(chaptersCount, 1, BibleCommon.Resources.Constants.AddParallelBibleTranslation);
                    result = SupplementalBibleManager.AddParallelBible(_oneNoteApp, selectedModuleInfo.ShortName, _logger);
                    _form.ExternalProcessingDone(BibleCommon.Resources.Constants.AddParallelBibleTranslationFinishMessage);
                }
                else
                {
                    int chaptersCount = ModulesManager.GetBibleChaptersCount(selectedModuleInfo.ShortName);
                    _form.PrepareForExternalProcessing(chaptersCount, 1, BibleCommon.Resources.Constants.CreateSupplementalBible);
                    _logger.Preffix = string.Format("{0} 1/2: ", BibleCommon.Resources.Constants.Stage);
                    SupplementalBibleManager.CreateSupplementalBible(_oneNoteApp, selectedModuleInfo.ShortName, _logger);

                    _form.PrepareForExternalProcessing(chaptersCount, 1, BibleCommon.Resources.Constants.LinkSupplementalBible);
                    _logger.Preffix = string.Format("{0} 2/2: ", BibleCommon.Resources.Constants.Stage);
                    result = SupplementalBibleManager.LinkSupplementalBibleWithMainBible(_oneNoteApp, 0, _logger);

                    _form.ExternalProcessingDone(BibleCommon.Resources.Constants.CreateSupplementalBibleFinishMessage);
                }

                if (result.Errors.Count > 0)
                {
                    var errorsForm = new BibleCommon.UI.Forms.ErrorsForm(result.Errors.ConvertAll(ex => ex.Message));
                    errorsForm.ShowDialog();
                }
            }
            catch (ProcessAbortedByUserException)
            {
                BibleCommon.Services.Logger.LogMessage("Process aborted by user");
                _form.ExternalProcessingDone(BibleCommon.Resources.Constants.ProcessAbortedByUser);
            }
            catch (Exception ex)
            {
                Logger.LogError(ex.ToString());
                _form.ExternalProcessingDone(BibleCommon.Resources.Constants.ErrorOccurred);
            }

            EnableUI(true);

            LoadUI();
        }

        private void EnableUI(bool enabled)
        {
            pnModules.Enabled = enabled;
            chkUseSupplementalBible.Enabled = enabled;
            btnOk.Enabled = enabled;
            btnCancel.Enabled = enabled;
        }

        private void LoadModules()
        {
            for (int i = 0; i < SettingsManager.Instance.SupplementalBibleModules.Count; i++)
            {
                AddModuleRow(_modules.First(m => m.ShortName == SettingsManager.Instance.SupplementalBibleModules[i]), i, _top);
                _top += 30;
            }

            _btnAddNewModule.Top = _top;
        }

        private void AddModuleRow(ModuleInfo moduleInfo, int index, int top)
        {
            Label lblName = new Label();
            lblName.Text = moduleInfo.Name;
            lblName.Top = top + 5;
            lblName.Left = 0;
            lblName.Width = 265;
            pnModules.Controls.Add(lblName);

            Button btnDel = new Button();
            btnDel.Image = BibleConfigurator.Properties.Resources.del;
            btnDel.Enabled = index == SettingsManager.Instance.SupplementalBibleModules.Count - 1;
            FormExtensions.SetToolTip(btnDel, BibleCommon.Resources.Constants.DeleteThisModule);
            btnDel.Tag = moduleInfo.ShortName;
            btnDel.Top = top;
            btnDel.Left = 268;
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
            if (MessageBox.Show(BibleCommon.Resources.Constants.DeleteThisModuleQuestion, BibleCommon.Resources.Constants.Warning, MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
                == System.Windows.Forms.DialogResult.Yes)
            {
                try
                {
                    int chaptersCount = ModulesManager.GetBibleChaptersCount(moduleName);
                    _form.PrepareForExternalProcessing(chaptersCount, 1, BibleCommon.Resources.Constants.RemoveParallelBibleTranslation);
                    SupplementalBibleManager.RemoveLastSupplementalBibleModule(_oneNoteApp, _logger);
                    _form.ExternalProcessingDone(BibleCommon.Resources.Constants.RemoveParallelBibleTranslationFinishMessage);
                }
                catch (ProcessAbortedByUserException)
                {
                    BibleCommon.Services.Logger.LogMessage("Process aborted by user");
                    _form.ExternalProcessingDone(BibleCommon.Resources.Constants.ProcessAbortedByUser);
                }
                catch (Exception ex)
                {
                    Logger.LogError(ex.ToString());
                    _form.ExternalProcessingDone(BibleCommon.Resources.Constants.ErrorOccurred);
                }

                LoadFormData();
                return true;
            }

            return false;
        }  


        private void SupplementalBibleForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            _oneNoteApp = null;
            _form = null;
            _logger.Dispose();
        }

        private void chkUseSupplementalBible_CheckedChanged(object sender, EventArgs e)
        {
            bool needToUpdate = true;

            if (_wasLoaded && !chkUseSupplementalBible.Checked 
                && !string.IsNullOrEmpty(SettingsManager.Instance.GetValidSupplementalBibleNotebookId(_oneNoteApp)))
            {
                if (MessageBox.Show(BibleCommon.Resources.Constants.DeleteSupplementalBibleQuestion,
                    BibleCommon.Resources.Constants.Warning, MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
                    == System.Windows.Forms.DialogResult.Yes)
                {
                    SupplementalBibleManager.RemoveSupplementalBible(_oneNoteApp);
                }
                else
                {
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

        private void AddModulesComboBox()
        {            
            _cbModule = new ComboBox();
            _cbModule.DropDownStyle = ComboBoxStyle.DropDownList;
            _cbModule.Width = 260;
            _cbModule.Top = _top;
            _cbModule.ValueMember = "Name";

            _top = _top + 30;

            foreach (var moduleInfo in _modules)
            {
                if (BibleParallelTranslationManager.IsModuleSupported(moduleInfo)
                    && !SettingsManager.Instance.SupplementalBibleModules.Contains(moduleInfo.ShortName))
                    _cbModule.Items.Add(moduleInfo);
            }

            if (_cbModule.Items.Count > 0)
                _cbModule.SelectedIndex = 0;

            pnModules.Controls.Add(_cbModule);
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            if (_needToCommitChanges)
                CommitChanges();

            this.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void SupplementalBibleForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            _logger.AbortedByUsers = true;            
        }

    }
}
