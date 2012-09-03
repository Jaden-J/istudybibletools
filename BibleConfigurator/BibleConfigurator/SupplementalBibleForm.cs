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

        public SupplementalBibleForm(Microsoft.Office.Interop.OneNote.Application oneNoteApp, MainForm form)
        {
            _oneNoteApp = oneNoteApp;
            _form = form;
            _modules = ModulesManager.GetModules();

            InitializeComponent();            
        }

        private void SupplementalBibleForm_Load(object sender, EventArgs e)
        {
            chkUseSupplementalBible.Checked = !string.IsNullOrEmpty(SettingsManager.Instance.GetValidSupplementalBibleNotebookId(_oneNoteApp));

            LoadModules();

            chkUseSupplementalBible_CheckedChanged(this, null);
        }

        private void LoadModules()
        {
            int top = 10;

            if (SettingsManager.Instance.SupplementalBibleModules.Count > 0)
            {
                for (int i = 0; i < SettingsManager.Instance.SupplementalBibleModules.Count; i++)
                {
                    AddModuleRow(SettingsManager.Instance.SupplementalBibleModules[i], i, top);
                    top += 30;
                }
            }
            else
            {
                AddModulesComboBox();
            }
        }

        private void AddModuleRow(string moduleName, int index, int top)
        {
            Label l = new Label();
            l.Text = moduleName;
            l.Top = top + 5;
            l.Left = 0;
            l.Width = 365;
            pnModules.Controls.Add(l);

            Button bDel = new Button();
            bDel.Image = BibleConfigurator.Properties.Resources.del;
            bDel.Enabled = index != 0 && index == SettingsManager.Instance.SupplementalBibleModules.Count - 1;
            SetToolTip(bDel, BibleCommon.Resources.Constants.DeleteThisModule);
            bDel.Tag = moduleName;
            bDel.Top = top;
            bDel.Left = 600;
            bDel.Width = bDel.Height;
            bDel.Click += new EventHandler(btnDeleteModule_Click);
            pnModules.Controls.Add(bDel);
        }

        void btnDeleteModule_Click(object sender, EventArgs e)
        {
            var btn = (Button)sender;
            var moduleName = (string)btn.Tag;

            DeleteModuleWithConfirm(moduleName);
        }

        private bool DeleteModuleWithConfirm(string moduleName)
        {
            if (SettingsManager.Instance.SupplementalBibleModules.Last() != moduleName)
                throw new NotSupportedException("Only last module can be deleted.");

            if (MessageBox.Show(BibleCommon.Resources.Constants.DeleteThisModuleQuestion, BibleCommon.Resources.Constants.Warning, MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
                == System.Windows.Forms.DialogResult.Yes)
            {
                SupplementalBibleManager.RemoveLastSupplementalModule();
                //ModulesManager.DeleteModule(moduleName);

                //ReLoadModulesInfo();
                return true;
            }

            return false;
        }


        private ToolTip _toolTip = null;
        private void SetToolTip(Control c, string toolTip)
        {
            if (_toolTip == null)
            {
                _toolTip = new ToolTip();

                _toolTip.AutoPopDelay = 5000;
                _toolTip.InitialDelay = 1000;
                _toolTip.ReshowDelay = 500;
                _toolTip.ShowAlways = true;
            }

            _toolTip.SetToolTip(c, toolTip);
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
            cb.Width = 300;

            foreach (var moduleInfo in _modules)
            {
                if (!SettingsManager.Instance.SupplementalBibleModules.Contains(moduleInfo.ShortName))
                    cb.Items.Add(moduleInfo.Name); 
            }

            pnModules.Controls.Add(cb);
        }
    }
}
