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
    public partial class RenameSectionGroupsForm : Form
    {
        public string SectionGroupName { get; private set; }

        public RenameSectionGroupsForm(string sectionGroupName)
        {
            InitializeComponent();
            SectionGroupName = sectionGroupName;
        }

        private void RenameSectionGroupsForm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                this.Close();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(tbSectionGroupName.Text))
            {
                SectionGroupName = tbSectionGroupName.Text;
                this.Close();
            }
        }

        private void RenameSectionGroupsForm_Load(object sender, EventArgs e)
        {
            try
            {
                tbSectionGroupName.Text = SectionGroupName;
            }
            catch (Exception ex)
            {
                FormLogger.LogError(ex);
            }
        }
    }
}
