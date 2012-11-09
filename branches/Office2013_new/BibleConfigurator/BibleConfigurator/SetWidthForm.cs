using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using BibleConfigurator.Properties;
using BibleCommon.Services;

namespace BibleConfigurator
{
    public partial class SetWidthForm : Form
    {
        public SetWidthForm()
        {
            InitializeComponent();
        }

        private int _biblePagesWidth;
        public int BiblePagesWidth
        {
            get
            {
                return _biblePagesWidth;
            }
        }

        private void SetWidthForm_Load(object sender, EventArgs e)
        {
            try
            {
                tbBiblePageWidth.Text = SettingsManager.Instance.PageWidth_Bible.ToString();
            }
            catch (Exception ex)
            {
                FormLogger.LogError(ex);
            }
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            if (int.TryParse(tbBiblePageWidth.Text, out _biblePagesWidth))
            {
                SettingsManager.Instance.PageWidth_Bible = _biblePagesWidth;
                SettingsManager.Instance.Save();

                this.DialogResult = System.Windows.Forms.DialogResult.OK;
                this.Close();
            }
        }
    }
}
