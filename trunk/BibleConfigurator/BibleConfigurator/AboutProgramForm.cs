using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using BibleCommon.Resources;
using System.Diagnostics;
using BibleCommon.Services;
using BibleCommon.Helpers;
using System.Threading;

namespace BibleConfigurator
{
    public partial class AboutProgramForm : Form
    {
        public AboutProgramForm()
        {
            this.SetFormUICulture();

            InitializeComponent();
        }

        private void AboutProgramForm_Load(object sender, EventArgs e)
        {
            try
            {
                hlSite.Text = Constants.WebSiteUrl;
                lblAuthor.Text = Constants.Author;
                lblVersion.Text = SettingsManager.Instance.CurrentVersion.ToString();

                string versionMessage = string.Format(" v{0}", SettingsManager.Instance.CurrentVersion);
                this.Text += versionMessage;                

                new Thread(CheckForNewerVersion).Start();
            }
            catch (Exception ex)
            {
                FormLogger.LogError(ex);
            }
        }

        public void CheckForNewerVersion()
        {
            var vsm = new VersionOnServerManager();
            if (vsm.NeedToUpdate())
            {
                FormExtensions.SetControlPropertyThreadSafe(lblNewVersion, "Visible", true);                
            }
        }


        private void hlSite_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start(Constants.WebSiteUrl);
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void lblNewVersion_Click(object sender, EventArgs e)
        {
            Process.Start(Utils.GetUpdateProgramWebSitePageUrl());
        }

        private bool _wasShown = false;
        private void AboutProgramForm_Shown(object sender, EventArgs e)
        {
            if (!_wasShown)
            {
                this.SetFocus();
                _wasShown = true;
            }
        }
    }
}
