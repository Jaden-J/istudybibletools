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

namespace BibleCommon.UI.Forms
{
    public partial class MessageForm : Form
    {
        public enum Severity
        {
            Custom,
            Information,
            Error            
        }

        public string Message { get; set; }
        public Severity MessageSeverity { get; set; }
        public string MessageCaption { get; set; }
        public MessageBoxButtons MessageButtons { get; set; }
        public MessageBoxIcon MessageIcon { get; set; }        

        public MessageForm()
        {
            InitializeComponent();            
        }

        public MessageForm(string message, Severity messageSeverity = Severity.Information)
            : this()
        {
            this.Message = message;
            this.MessageSeverity = messageSeverity;
        }

        public MessageForm(string message, string caption, MessageBoxButtons buttons, MessageBoxIcon icon)
            : this()
        {
            this.Message = message;
            this.MessageCaption = caption;
            this.MessageButtons = buttons;
            this.MessageIcon = icon;
            this.MessageSeverity = Severity.Custom;
        }

        private void MessageForm_Load(object sender, EventArgs e)
        {
            this.SetFocus();
            this.Top = -50;
            switch (MessageSeverity)
            {
                case Severity.Information:
                    FormLogger.LogMessage(Message);
                    break;                
                case Severity.Error:
                    FormLogger.LogError(Message);
                    break;
                case Severity.Custom:
                    this.DialogResult = MessageBox.Show(Message, MessageCaption, MessageButtons, MessageIcon);
                    break;
            }
            this.Close();
        }        
    }
}
