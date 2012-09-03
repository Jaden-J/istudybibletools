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

        public SupplementalBibleForm(Microsoft.Office.Interop.OneNote.Application oneNoteApp, MainForm form)
        {
            _oneNoteApp = oneNoteApp;
            _form = form;
            _modules = ModulesManager.GetModules();
            _top = 10;

            InitializeComponent();            
        }

        private void SupplementalBibleForm_Load(object sender, EventArgs e)
        {
            chkUseSupplementalBible.Checked = !string.IsNullOrEmpty(SettingsManager.Instance.GetValidSupplementalBibleNotebookId(_oneNoteApp));

            GenerateNewModuleButton();            

            LoadModules();

            chkUseSupplementalBible_CheckedChanged(this, null);
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
            AddModulesComboBox();

            _btnAddNewModule.Top = _top;
        }

        private void LoadModules()
        {
            if (SettingsManager.Instance.SupplementalBibleModules.Count > 0)
            {
                for (int i = 0; i < SettingsManager.Instance.SupplementalBibleModules.Count; i++)
                {
                    AddModuleRow(_modules.First(m => m.ShortName == SettingsManager.Instance.SupplementalBibleModules[i]).Name, i, _top);
                    _top += 30;
                }
            }
            else
            {
                AddModulesComboBox();
            }

            _btnAddNewModule.Top = _top;
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
            FormExtensions.SetToolTip(bDel, BibleCommon.Resources.Constants.DeleteThisModule);
            bDel.Tag = moduleName;
            bDel.Top = top;
            bDel.Left = 250;
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
            cb.Width = 250;
            cb.Top = _top;

            _top = _top + 30;

            foreach (var moduleInfo in _modules)
            {
                if (!SettingsManager.Instance.SupplementalBibleModules.Contains(moduleInfo.ShortName))
                    cb.Items.Add(moduleInfo.Name); 
            }

            pnModules.Controls.Add(cb);
        }
    }
}
