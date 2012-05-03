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
using System.IO;

namespace BibleConfigurator
{
    public partial class AboutModuleForm : Form
    {
        public string ModuleName { get; set; }

        public AboutModuleForm(string moduleName)
        {
            this.SetFormUICulture();

            InitializeComponent();

            this.ModuleName = moduleName;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void AboutModule_Load(object sender, EventArgs e)
        {
            ModuleInfo module = ModulesManager.GetModuleInfo(ModuleName);

            this.Text = lblTitle.Text = module.Name;
            lblLocation.Text = Path.Combine(ModulesManager.GetModulesDirectory(), ModuleName);
        }
    }
}
