using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using BibleCommon.Services;

namespace BibleConfigurator
{
    public partial class ParallelBibleChecker : Form
    {
        public ParallelBibleChecker()
        {
            InitializeComponent();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void ParallelBibleChecker_Load(object sender, EventArgs e)
        {
            SetDataSource(cbBaseModule);  
            SetDataSource(cbParallelModule);            
        }

        private void SetDataSource(ComboBox cb)
        {
            cb.DataSource = ModulesManager.GetModules(); // приходится каждый раз загружать, чтобы разные были дата сорсы - иначе они вместе меняются
            cb.DisplayMember = "ShortName";
            cb.ValueMember = "ShortName";
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            string baseModule = (string)cbBaseModule.SelectedValue;
            string parallelModule = (string)cbParallelModule.SelectedValue;

            var oneNoteApp = new Microsoft.Office.Interop.OneNote.Application();
            var manager = new BibleParallelTranslationManager(oneNoteApp, baseModule, parallelModule, SettingsManager.Instance.NotebookId_Bible);
            manager.ForCheckOnly = true;
            var result = manager.IterateBaseBible(null, false, true, null);
            if (result.Errors.Count > 0)
            {
                var errorsForm = new BibleCommon.UI.Forms.ErrorsForm(result.Errors.ConvertAll(ex => ex.Message));
                errorsForm.ErrorsDecription = string.Format("{0} -> {1}", baseModule, parallelModule);
                errorsForm.ShowDialog();
            }
            else
                MessageBox.Show("There is no errors");
        }
    }
}
