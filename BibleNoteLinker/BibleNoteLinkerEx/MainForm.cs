using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using BibleNoteLinkerEx.Properties;
using System.Configuration;
using BibleCommon;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using BibleCommon.Helpers;
using BibleCommon.Services;
using System.Xml.Linq;
using BibleCommon.Consts;
using BibleCommon.Common;

namespace BibleNoteLinkerEx
{
    public partial class MainForm : Form
    {
        private Microsoft.Office.Interop.OneNote.Application _oneNoteApp;
        

        public MainForm()
        {
            InitializeComponent();
            _oneNoteApp = new Microsoft.Office.Interop.OneNote.Application();
        }


        private int _originalFormHeight;
        const int FirstFormHeight = 185;
        const int SecondFormHeight = 250;
        private bool _processAbortedByUser;

        private delegate void SetControlPropertyThreadSafeDelegate(Control control, string propertyName, object propertyValue);

        public static void SetControlPropertyThreadSafe(Control control, string propertyName, object propertyValue)
        {
            if (control.InvokeRequired)
            {
                control.Invoke(new SetControlPropertyThreadSafeDelegate(SetControlPropertyThreadSafe), new object[] { control, propertyName, propertyValue });
            }
            else
            {
                control.GetType().InvokeMember(propertyName, BindingFlags.SetProperty, null, control, new object[] { propertyValue });
            }
        }

        [DllImport("user32.dll")]
        static extern bool SetForegroundWindow(IntPtr hWnd);


        private void btnOk_Click(object sender, EventArgs e)
        {
            BibleNoteLinkerEx.Properties.Settings.Default.AllPages = rbAnalyzeAllPages.Checked;
            BibleNoteLinkerEx.Properties.Settings.Default.Changed = rbAnalyzeChangedPages.Checked;
            BibleNoteLinkerEx.Properties.Settings.Default.Force = chkForce.Checked;
            BibleNoteLinkerEx.Properties.Settings.Default.Save();            

            this.Height = SecondFormHeight;
            EnableBaseElements(false);

            BibleCommon.Services.Logger.Init("BibleNoteLinkerEx");
            BibleCommon.Services.Logger.SetOutputListBox(lbLog);

            try
            {
                StartAnalyze();
            }
            catch (ProcessAbortedByUserException)
            {
                Logger.LogMessage(string.Empty, false, true, false);
                Logger.LogMessage("Операция прервана пользователем.");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                BibleCommon.Services.Logger.LogError(ex);
            }
            finally
            {
                EnableBaseElements(true);
                BibleCommon.Services.Logger.Done();
            }

            pbMain.Value = pbMain.Maximum = 1;

            if (!Logger.ErrorWasLogged)
                LogHighLevelMessage("Успешно завершено.", null, null);
            else
                LogHighLevelMessage("Завершено с ошибками.", null, null);
        }    
       

        private void EnableBaseElements(bool enabled)
        {
            panel1.Enabled = enabled;
            tsmiSeelctNotebooks.Enabled = enabled;
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
            if (BibleNoteLinkerEx.Properties.Settings.Default.AllPages)
                rbAnalyzeAllPages.Checked = true;
            else if (BibleNoteLinkerEx.Properties.Settings.Default.Changed)
                rbAnalyzeChangedPages.Checked = true;

            if (BibleNoteLinkerEx.Properties.Settings.Default.Force)
                chkForce.Checked = true;

            lblInfo.Visible = false;
            _originalFormHeight = this.Height;
            this.Height = FirstFormHeight;

            new Thread(CheckForNewerVersion).Start();
        }       

        public void CheckForNewerVersion()
        {
            if (VersionOnServerManager.NeedToUpdate())
            {
                SetControlPropertyThreadSafe(lblInfo, "Text",
@"Доступна новая версия программы
на сайте http://IStudyBibleTools.ru. 
Кликните, чтобы перейти на страницу загрузки.");

                SetControlPropertyThreadSafe(this, "Size", new Size(this.Size.Width, this.Size.Height + 50));
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
            Process.Start(BibleCommon.Consts.Constants.DownloadPageUrl);
        }

        private bool _detailsWereShown = false;
        private void llblDetails_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (!_detailsWereShown)
            {
                llblDetails.Text = "Скрыть детали";
                this.Height = _originalFormHeight;                
            }
            else
            {
                llblDetails.Text = "Показать детали";
                this.Height = SecondFormHeight;                
            }

            _detailsWereShown = !_detailsWereShown;
        }

        private void tsmiSeelctNotebooks_Click(object sender, EventArgs e)
        {
            SelectNoteBooks form = new SelectNoteBooks(_oneNoteApp);
            form.ShowDialog();
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            _processAbortedByUser = true;
        }
    }
}
