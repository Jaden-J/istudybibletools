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
            InitializeComponent();
        }

        private void AboutProgramForm_Load(object sender, EventArgs e)
        {
            hlSite.Text = Constants.SiteUrl;            
            lblAuthor.Text = Constants.Author;

            new Thread(CheckForNewerVersion).Start();
        }

        public void CheckForNewerVersion()
        {
            if (VersionOnServerManager.NeedToUpdate())
            {
                FormExtensions.SetControlPropertyThreadSafe(lblNewVersion, "Visible", true);                
            }
        }


        private void hlSite_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start(Constants.SiteUrl);
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void lblNewVersion_Click(object sender, EventArgs e)
        {
            Process.Start(BibleCommon.Consts.Constants.DownloadPageUrl);
        }
    }
}
