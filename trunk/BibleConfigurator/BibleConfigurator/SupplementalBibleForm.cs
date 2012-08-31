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

namespace BibleConfigurator
{
    public partial class SupplementalBibleForm : Form
    {
        private Microsoft.Office.Interop.OneNote.Application _oneNoteApp;
        private MainForm _form;

        public SupplementalBibleForm(Microsoft.Office.Interop.OneNote.Application oneNoteApp, MainForm form)
        {
            _oneNoteApp = oneNoteApp;
            _form = form;

            InitializeComponent();            
        }

        private void SupplementalBibleForm_Load(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(SettingsManager.Instance.NotebookId_SupplementalBible))
                if (!OneNoteUtils.NotebookExists(_oneNoteApp, SettingsManager.Instance.NotebookId_SupplementalBible))
                {
                    SettingsManager.Instance.NotebookId_SupplementalBible = null;
                    SettingsManager.Instance.Save();
                }

            chkUseSupplementalBible.Checked = !string.IsNullOrEmpty(SettingsManager.Instance.NotebookId_SupplementalBible);

            LoadModules();
        }

        private void LoadModules()
        {
            if (SettingsManager.Instance.SupplementalBibleModules.Count > 0)
            {
                foreach (var moduleName in SettingsManager.Instance.SupplementalBibleModules)
                {

                }
            }
            else
            {
                AddModulesComboBox();
            }
        }

        private void SupplementalBibleForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            _oneNoteApp = null;
        }

        private void chkUseSupplementalBible_CheckedChanged(object sender, EventArgs e)
        {
            pnModules.Enabled = chkUseSupplementalBible.Checked;
        }

        private void SupplementalBibleForm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                Close();
        }

        private void AddModulesComboBox()
        {
            ComboBox cb = new ComboBox();
            foreach (var moduleName in ModulesManager.GetModules())
            {
                cb.Items.Add(moduleName);
            }
            pnModules.Controls.Add(cb);
        }
    }
}
