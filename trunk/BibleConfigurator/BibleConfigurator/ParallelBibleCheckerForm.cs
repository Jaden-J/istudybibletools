using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using BibleCommon.Services;
using BibleCommon.Common;
using BibleCommon.Helpers;
using BibleCommon.UI.Forms;

namespace BibleConfigurator
{
    public partial class ParallelBibleCheckerForm : Form
    {
        public string ModuleToCheckName { get; set; }
        public bool AutoStart { get; set; }

        private MainForm _mainForm;        
        private CustomFormLogger _formLogger;
        private ErrorsForm _errorsForm;

        public ParallelBibleCheckerForm(MainForm mainForm)
        {
            InitializeComponent();            
            _mainForm = mainForm;
            _formLogger = new CustomFormLogger(_mainForm);
            _formLogger.Preffix = "Checking: ";
            _errorsForm = new ErrorsForm();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void ParallelBibleChecker_Load(object sender, EventArgs e)
        {
            try
            {
                SetDataSource(cbBaseModule);
                SetDataSource(cbParallelModule);

                LoadControlsState();

                if (AutoStart)
                    btnOk_Click(btnOk, null);
            }
            catch (Exception ex)
            {
                FormLogger.LogError(ex);
            }
        }

        private void LoadControlsState()
        {
            if (!string.IsNullOrEmpty(ModuleToCheckName))
            {
                cbBaseModule.SelectedValue = ModuleToCheckName;
                chkWithAllModules.Checked = true;
                cbParallelModule.Enabled = false;
            }
        }

        private void CloseResources()
        {
            _mainForm = null;
            _errorsForm.Dispose();
            _errorsForm = null;
            _formLogger.Dispose();
        }

        private void SetDataSource(ComboBox cb)
        {
            cb.DataSource = GetModules();   // приходится каждый раз загружать, чтобы разные были дата сорсы - иначе они вместе меняются
            cb.DisplayMember = "ShortName";
            cb.ValueMember = "ShortName";
        }

        private List<ModuleInfo> GetModules()
        {
            return ModulesManager.GetModules(true).Where(m => m.Type == ModuleType.Bible || m.Type == ModuleType.Strong).ToList(); 
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            string baseModule = (string)cbBaseModule.SelectedValue;
            string parallelModule = (string)cbParallelModule.SelectedValue;
            
            FormExtensions.EnableAll(false, this.Controls, btnClose);
            this.SetFocus();
            System.Windows.Forms.Application.DoEvents();

            _errorsForm.ClearErrors();

            var allModules = GetModules();

            try
            {
                if (rbCheckOneModule.Checked)
                {
                    if (!chkWithAllModules.Checked)
                    {
                        CheckModule(baseModule, parallelModule);
                    }
                    else
                    {
                        _mainForm.PrepareForExternalProcessing((allModules.Count - 1) * 2, 1, "Start checking");

                        foreach (var pModule in allModules.Where(m => m.ShortName != baseModule))
                        {
                            _formLogger.LogMessage("{0} -> {1}", baseModule, pModule.ShortName);
                            CheckModule(baseModule, pModule.ShortName);

                            _formLogger.LogMessage("{0} -> {1}", pModule.ShortName, baseModule);
                            CheckModule(pModule.ShortName, baseModule);
                        }

                        _mainForm.ExternalProcessingDone("Checking complete");
                    }
                }
                else
                {
                    _mainForm.PrepareForExternalProcessing(allModules.Count * (allModules.Count - 1), 1, "Start checking");

                    foreach (var bModule in allModules)
                    {
                        foreach (var pModule in allModules.Where(m => m.ShortName != bModule.ShortName))
                        {
                            _formLogger.LogMessage("{0} -> {1}", bModule.ShortName, pModule.ShortName);
                            CheckModule(bModule.ShortName, pModule.ShortName);
                        }
                    }

                    _mainForm.ExternalProcessingDone("Checking complete");
                }

                if (_errorsForm.AllErrors.Any(errors => errors.Count > 0))
                    _errorsForm.ShowDialog();
                else
                    MessageBox.Show("There is no errors");

                if (AutoStart)
                    Close();
                else
                {
                    FormExtensions.EnableAll(true, this.Controls);
                    LoadControlsState();
                }
            }
            catch (ProcessAbortedByUserException)
            {
                _mainForm.ExternalProcessingDone(BibleCommon.Resources.Constants.ProcessAbortedByUser);
            }
        }

        private void CheckModule(string primaryModuleName, string parallelModuleName)
        {
            var manager = new BibleParallelTranslationManager(null, primaryModuleName, parallelModuleName, SettingsManager.Instance.NotebookId_Bible);
            manager.ForCheckOnly = true;
            var result = manager.IterateBaseBible(null, false, true, null, true);

            if (result.Errors.Count > 0)
            {
                _errorsForm.AllErrors.Add(
                    new ErrorsForm.ErrorsList(result.Errors.ConvertAll(ex => ex.Message))
                    {
                        ErrorsDecription = string.Format("{0} -> {1}", primaryModuleName, parallelModuleName)
                    });
            }            
        }

        private void chkWithAllModules_CheckedChanged(object sender, EventArgs e)
        {
            cbParallelModule.Enabled = !((CheckBox)sender).Checked;
        }

        private void rbCheckAllModules_CheckedChanged(object sender, EventArgs e)
        {
            cbBaseModule.Enabled = !((RadioButton)sender).Checked;
            cbParallelModule.Enabled = !((RadioButton)sender).Checked && !chkWithAllModules.Checked;
            chkWithAllModules.Enabled = !((RadioButton)sender).Checked;
        }

        private void ParallelBibleCheckerForm_FormClosing(object sender, FormClosingEventArgs e)
        {   
            _formLogger.AbortedByUsers = true;
        }

        private void ParallelBibleCheckerForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            CloseResources();
        }     
    }
}
