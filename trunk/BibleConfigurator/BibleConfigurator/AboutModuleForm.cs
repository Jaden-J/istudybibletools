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

            this.Text = lblTitle.Text = string.Format("{0} ({1} {2})", module.Name, module.BibleStructure.BibleBooks.Count, BibleCommon.Resources.Constants.Books);
            lblLocation.Text = Path.Combine(ModulesManager.GetModulesDirectory(), ModuleName);

            int top = 5;

            foreach (var book in module.BibleStructure.BibleBooks)
            {
                var lblBook = new Label();
                lblBook.Top = top;
                lblBook.Left = 0;
                lblBook.Text = string.Format("{0}:", book.Name);
                lblBook.Font = new System.Drawing.Font("Microsoft Sans Serif", (float)8.25, FontStyle.Bold);
                lblBook.Width = GetLabelWidth(lblBook);
                pnBooks.Controls.Add(lblBook);                

                var lblAbbr = new Label();
                lblAbbr.Top = top;
                lblAbbr.Left = lblBook.Width + 5;
                lblAbbr.Text = string.Join("  ", book.Abbreviations.ToArray());
                lblAbbr.Width = GetLabelWidth(lblAbbr);
                pnBooks.Controls.Add(lblAbbr);  

                top += 25;
            }
        }

        private int GetLabelWidth(Label lbl)
        {
            return (int)lbl.CreateGraphics().MeasureString(lbl.Text, lbl.Font).Width + 20;
        }
    }
}
