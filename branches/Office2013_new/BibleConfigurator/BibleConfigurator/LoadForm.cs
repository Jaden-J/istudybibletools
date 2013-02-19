using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using BibleCommon.Helpers;
using BibleCommon.Services;

namespace BibleConfigurator
{
    public partial class LoadForm : Form
    {        

        // todo: надо бы добавить картинку на эту форму, но учитывать, что эта форма используется в двух случаях:
            // - при загрузке BibleConfigurator
            // - при загрузке модуля
        public LoadForm()
        {
            InitializeComponent();            
        }

        private void LoadForm_Load(object sender, EventArgs e)
        {
            try
            {
                pbImage.Top = (this.Height - pbImage.Height) / 2;
                pbImage.Left = (this.Width - pbImage.Width) / 2;
            }
            catch (Exception ex)
            {
                FormLogger.LogError(ex);
            }
        }
    }
}
