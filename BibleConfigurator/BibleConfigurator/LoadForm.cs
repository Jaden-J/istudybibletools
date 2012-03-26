using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace BibleConfigurator
{
    public partial class LoadForm : Form
    {
        public LoadForm()
        {
            InitializeComponent();
        }

        private void LoadForm_Load(object sender, EventArgs e)
        {
            pbImage.Top = (this.Height - pbImage.Height) / 2;
            pbImage.Left = (this.Width - pbImage.Width) / 2;
        }
    }
}
