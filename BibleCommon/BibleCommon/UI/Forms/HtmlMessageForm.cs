using BibleCommon.Services;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using BibleCommon.Helpers;

namespace BibleCommon.UI.Forms
{
    public partial class HtmlMessageForm : Form
    {
        public int MessageCode { get; set; }

        public HtmlMessageForm()
        {
            InitializeComponent();
        }

        public HtmlMessageForm(int messageCode, string message, string caption)
            : this()
        {
            this.MessageCode = messageCode;
            this.Text = caption;
            this.webBrowser.DocumentText = message;            
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void chkDontShow_CheckedChanged(object sender, EventArgs e)
        {
            if (chkDontShow.Checked)
                ShownMessagesManager.SetMessageWasShown(MessageCode);
            else
                ShownMessagesManager.ClearMessageWasShown(MessageCode);

            SettingsManager.Instance.Save();
        }

        private bool _wasShown = false;
        private void HtmlMessageForm_Shown(object sender, EventArgs e)
        {
            if (!_wasShown)
            {
                this.SetFocus();
                _wasShown = true;
            }
        }
    }
}
