using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using BibleCommon.Services;
using BibleCommon.Common;
using System.Reflection;
using System.Diagnostics;
using System.Threading;
using BibleCommon.Helpers;

namespace BibleNoteLinker
{
    public partial class MainForm : Form
    {
        private Microsoft.Office.Interop.OneNote.Application _oneNoteApp;

        public MainForm()
        {
            this.SetFormUICulture();

            InitializeComponent();
            _oneNoteApp = OneNoteUtils.CreateOneNoteAppSafe();
        }

        private int _originalFormHeight;        
        const int FirstFormHeight = 192;
        const int SecondFormHeight = 255;
        private bool _processAbortedByUser;
        private bool _wasStartAnalyze = false;
        private bool _wasAnalyzed = false;

        private void btnOk_Click(object sender, EventArgs e)
        {
            btnOk.Enabled = false;
            Application.DoEvents();

            if (_wasAnalyzed)
            {
                this.Close();
                return;
            }

            if (!SettingsManager.Instance.IsConfigured(ref _oneNoteApp))
            {
                MessageBox.Show(BibleCommon.Resources.Constants.Error_SystemIsNotConfigured);
                this.Close();
                return;
            }

            _wasStartAnalyze = true;

            BibleNoteLinker.Properties.Settings.Default.AllPages = rbAnalyzeAllPages.Checked;
            BibleNoteLinker.Properties.Settings.Default.Changed = rbAnalyzeChangedPages.Checked;
            BibleNoteLinker.Properties.Settings.Default.Force = chkForce.Checked;
            BibleNoteLinker.Properties.Settings.Default.Save();

            try
            {
                PrepareForAnalyze();

                DateTime dt = DateTime.Now;
                Logger.LogMessageParams("{0}: {1}", BibleCommon.Resources.Constants.StartTime, dt.ToLongTimeString());
                StartAnalyze();
                Logger.LogMessageParams("{0}: {1}", BibleCommon.Resources.Constants.TimeSpent, DateTime.Now.Subtract(dt));
                this.SetFocus();                
            }
            catch (ProcessAbortedByUserException)
            {
                Logger.LogMessageParams(BibleCommon.Resources.Constants.ProcessAbortedByUser);
            }
            catch (Exception ex)
            {
                FormLogger.LogError(ex);                
            }

            pbMain.Value = pbMain.Maximum = 1;

            string message;            

            if (Logger.ErrorWasLogged)
            {
                message = BibleCommon.Resources.Constants.FinishWithErrors;
                llblShowErrors.Visible = true;
                llblShowErrors.Text = BibleCommon.Resources.Constants.ShowErrors;
            }
            else if (Logger.WarningWasLogged)
            {
                message = BibleCommon.Resources.Constants.FinishWithWarnings;
                llblShowErrors.Visible = true;
                llblShowErrors.Text = BibleCommon.Resources.Constants.ShowWarnings;
            }
            else
            {
                message = BibleCommon.Resources.Constants.FinishSuccessfully;
            }

            LogHighLevelMessage(message, null, null);
            Logger.LogMessageParams(message);

            btnOk.Text = BibleCommon.Resources.Constants.Close;
            btnOk.Enabled = true;
            _wasAnalyzed = true;
            Logger.Done();
        }

        private void PrepareForAnalyze()
        {
            lbLog.Items.Clear();
            lbLog.HorizontalExtent = 0;

            Logger.Init("BibleNoteLinker");
            Logger.SetOutputListBox(lbLog);

            if (!_detailsWereShown)
                this.Height = SecondFormHeight;

            EnableControls(false);
            this.TopMost = false;

            llblShowErrors.Visible = false;
            LogHighLevelMessage(BibleCommon.Resources.Constants.NoteLinkerInitialization, null, null);            
        }


        private void EnableControls(bool enabled)
        {
            pbBaseElements.Enabled = enabled;
            tsmiSeelctNotebooks.Enabled = enabled;
            btnOk.Enabled = enabled;
        }


        private void MainForm_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.Escape:
                    this.Close();
                    break;
                case Keys.Space:
                    if (chkForce.Enabled)
                        chkForce.Checked = !chkForce.Checked;
                    e.SuppressKeyPress = true;
                    break;
            }
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            try
            {
                if (BibleNoteLinker.Properties.Settings.Default.AllPages)
                    rbAnalyzeAllPages.Checked = true;
                else if (BibleNoteLinker.Properties.Settings.Default.Changed)
                    rbAnalyzeChangedPages.Checked = true;

                if (BibleNoteLinker.Properties.Settings.Default.Force)
                    chkForce.Checked = true;

                lblInfo.Visible = false;
                _originalFormHeight = this.Height;
                this.Height = FirstFormHeight;

                if (SettingsManager.Instance.IsSingleNotebook)
                    tsmiSeelctNotebooks.Enabled = false;

                new Thread(CheckForNewerVersion).Start();
            }
            catch (Exception ex)
            {
                FormLogger.LogError(ex);
            }
        }

        public void CheckForNewerVersion()
        {
            if (VersionOnServerManager.NeedToUpdate())
            {
                FormExtensions.SetControlPropertyThreadSafe(lblInfo, "Visible", true);           
            }
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

        private void lblInfo_Click(object sender, EventArgs e)
        {
            Process.Start(BibleCommon.Resources.Constants.DownloadPageUrl);
        }

        private bool _detailsWereShown = false;
        private void llblDetails_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (!_detailsWereShown)
                ShowDetails();
            else
                HideDetails();
        }

        private void ShowDetails()
        {
            llblDetails.Text = BibleCommon.Resources.Constants.NoteLinkerHideDetails;
            this.Height = _originalFormHeight;
            _detailsWereShown = true;
        }

        private void HideDetails()
        {
            llblDetails.Text = BibleCommon.Resources.Constants.NoteLinkerShowDetails;
            this.Height = SecondFormHeight;
            _detailsWereShown = false;
        }

        private void tsmiSeelctNotebooks_Click(object sender, EventArgs e)
        {
            using (var form = new SelectNoteBooksForm())
            {
                form.ShowDialog();
            }
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (_wasStartAnalyze && !_wasAnalyzed)
            {
                var caption = BibleCommon.Resources.Constants.NoteLinkerFormCaptionOnClosing;
                string body;

                if (SettingsManager.Instance.StoreNotesPagesInFolder)
                {
                    body = BibleCommon.Resources.Constants.StopNotesAnalysis; 
                }
                else
                {
                    body = BibleCommon.Resources.Constants.NoteLinkerQuestionOnClosing;
                }

                if (MessageBox.Show(body, caption, MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == System.Windows.Forms.DialogResult.No)
                {
                    e.Cancel = true;
                }
                else
                {
                    _processAbortedByUser = true;
                }
            }
        }


        private void llblShowErrors_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            var messages = new List<string>();   
            messages.AddRange(Logger.Errors);              // если не было ошибок, то и так будет пустой список
            messages.AddRange(Logger.Warnings);

            messages = messages.Distinct().ToList();

            using (var errorsForm = new BibleCommon.UI.Forms.ErrorsForm(messages))
            {
                errorsForm.LogFilePath = Logger.LogFilePath;
                errorsForm.ShowDialog();
            }
        }

        private void MainForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            OneNoteUtils.ReleaseOneNoteApp(ref _oneNoteApp);
        }
    }
}
