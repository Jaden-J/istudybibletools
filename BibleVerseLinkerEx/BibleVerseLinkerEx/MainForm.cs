using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using System.Configuration;
using System.Runtime.InteropServices;
using System.Threading;
using BibleCommon;
using BibleCommon.Helpers;
using BibleCommon.Services;
using BibleCommon.Consts;

namespace BibleVerseLinkerEx
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            this.SetFormUICulture();

            InitializeComponent();
        }        

        private void btnOk_Click(object sender, EventArgs e)
        {
            btnOk.Enabled = false;
            Application.DoEvents();

            var oneNoteApp = OneNoteUtils.CreateOneNoteAppSafe();

            FormLogger.Initialize();
            BibleCommon.Services.Logger.Init("BibleVerseLinkerEx");

            try
            {
                if (!SettingsManager.Instance.IsConfigured(ref oneNoteApp))
                {
                    FormLogger.LogError(BibleCommon.Resources.Constants.Error_SystemIsNotConfigured);
                }
                else
                {
                    try
                    {
                        using (VerseLinker vlManager = new VerseLinker())
                        {
                            if (!string.IsNullOrEmpty(tbPageName.Text))
                                vlManager.DescriptionPageName = tbPageName.Text;
                            else
                            {
                                tbPageName.Text = SettingsManager.Instance.PageName_DefaultComments;
                                Application.DoEvents();
                            }

                            vlManager.Do();

                            if (!FormLogger.WasErrorLogged)
                            {
                                OneNoteUtils.SetActiveCurrentWindow(ref oneNoteApp);
                                this.Visible = false;
                                vlManager.SortCommentsPages();
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        FormLogger.LogError(OneNoteUtils.ParseErrorAndMakeItMoreUserFriendly(ex.Message));
                    }
                }

                btnOk.Enabled = true;

                if (!FormLogger.WasErrorLogged)
                {
                    this.Visible = false;
                    Properties.Settings.Default.LastPageName = tbPageName.Text;
                    Properties.Settings.Default.Save();
                    this.Close();
                }
            }
            finally
            {
                OneNoteUtils.ReleaseOneNoteApp(ref oneNoteApp);
                BibleCommon.Services.Logger.Done();
            }
        }
       

        private void MainForm_Load(object sender, EventArgs e)
        {
            try
            {
                tbPageName.Text = !string.IsNullOrEmpty(Properties.Settings.Default.LastPageName) 
                    ? Properties.Settings.Default.LastPageName
                    : SettingsManager.Instance.PageName_DefaultComments;
            }
            catch (Exception ex)
            {
                FormLogger.LogError(ex);
            }
        }

        private void MainForm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                this.Close();
        }

        private bool _wasShown = false;
        private void MainForm_Shown(object sender, EventArgs e)
        {
            if (!_wasShown)
            {
                this.SetFocus();
                _wasShown = true;
            }
        }                 
    }
}
