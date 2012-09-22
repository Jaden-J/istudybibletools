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
            var modules = ModulesManager.GetModules();

            cbBaseModule.DataSource = modules;
            cbBaseModule.DisplayMember = "ShortName";
            cbBaseModule.ValueMember = "ShortName";

            cbParallelModule.DataSource = modules;
            cbParallelModule.DisplayMember = "ShortName";
            cbParallelModule.ValueMember = "ShortName";
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            var oneNoteApp = new Microsoft.Office.Interop.OneNote.Application();
            var manager = new BibleParallelTranslationManager(oneNoteApp, (string)cbBaseModule.SelectedValue, (string)cbParallelModule.SelectedValue, SettingsManager.Instance.NotebookId_Bible);
            manager.ForCheckOnly = true; здесь
        }
    }
}
