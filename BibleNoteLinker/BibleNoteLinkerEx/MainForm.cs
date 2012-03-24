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


        private bool _wasStartAnalyze = false;
        private bool _wasAnalyzed = false;
        private void btnOk_Click(object sender, EventArgs e)
        {
            if (_wasAnalyzed)
            {
                this.Close();
                return;
            }

            _wasStartAnalyze = true;

            BibleNoteLinkerEx.Properties.Settings.Default.AllPages = rbAnalyzeAllPages.Checked;
            BibleNoteLinkerEx.Properties.Settings.Default.Changed = rbAnalyzeChangedPages.Checked;
            BibleNoteLinkerEx.Properties.Settings.Default.Force = chkForce.Checked;
            BibleNoteLinkerEx.Properties.Settings.Default.Save();

            try
            {
                PrepareForAnalyze();

                StartAnalyze();


            }
            catch (ProcessAbortedByUserException)
            {
                Logger.LogMessage("Операция прервана пользователем.");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Logger.LogError(ex);
            }

            pbMain.Value = pbMain.Maximum = 1;


            string message;
            if (!Logger.ErrorWasLogged)
                message = "Успешно завершено.";
            else
            {
                message = "Завершено с ошибками.";
                llblShowErrors.Visible = true;
            }

            LogHighLevelMessage(message, null, null);
            Logger.LogMessage(message);

            btnOk.Text = "Закрыть";
            btnOk.Enabled = true;
            _wasAnalyzed = true;
            Logger.Done();
        }

        private void PrepareForAnalyze()
        {
            lbLog.Items.Clear();
            lbLog.HorizontalExtent = 0;

            Logger.Init("BibleNoteLinkerEx");
            Logger.SetOutputListBox(lbLog);

            if (!_detailsWereShown)
                this.Height = SecondFormHeight;

            EnableControls(false);

            llblShowErrors.Visible = false;
            LogHighLevelMessage("Инициализация...", null, null);
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
            if (BibleNoteLinkerEx.Properties.Settings.Default.AllPages)
                rbAnalyzeAllPages.Checked = true;
            else if (BibleNoteLinkerEx.Properties.Settings.Default.Changed)
                rbAnalyzeChangedPages.Checked = true;

            if (BibleNoteLinkerEx.Properties.Settings.Default.Force)
                chkForce.Checked = true;

            lblInfo.Text = string.Empty;
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
                ShowDetails();
            else
                HideDetails();
        }

        private void ShowDetails()
        {
            llblDetails.Text = "Скрыть детали";
            this.Height = _originalFormHeight;
            _detailsWereShown = true;
        }

        private void HideDetails()
        {
            llblDetails.Text = "Показать детали";
            this.Height = SecondFormHeight;
            _detailsWereShown = false;
        }

        private void tsmiSeelctNotebooks_Click(object sender, EventArgs e)
        {
            SelectNoteBooksForm form = new SelectNoteBooksForm(_oneNoteApp);
            form.ShowDialog();
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (_wasStartAnalyze && !_wasAnalyzed)
            {
                if (MessageBox.Show("Вы действительно хотите прекратить работу программы? В некоторых случаях это может привести к неправильной сортировке страниц 'Сводные заметок', что решается только удалением всех страниц 'Сводные земеток'.",
                    "Закрыть программу?", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == System.Windows.Forms.DialogResult.No)
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
            ErrorsForm errors = new ErrorsForm();
            errors.ShowDialog();
        }
    }
}
