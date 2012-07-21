using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using BibleConfigurator.Properties;

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
            tbBiblePageWidth.Text = Settings.Default.BiblePagesWidth.ToString();
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            if (int.TryParse(tbBiblePageWidth.Text, out _biblePagesWidth))
            {
                Settings.Default.BiblePagesWidth = _biblePagesWidth;
                Settings.Default.Save();

                this.DialogResult = System.Windows.Forms.DialogResult.OK;
                this.Close();
            }
        }
    }
}
