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

namespace BibleNoteLinkerEx
{
    public partial class MainForm : Form
    {
        const string Arg_AllPages = "-allpages";        
        const string Arg_Changed = "-changed";
        const string Arg_Force = "-force";

        public MainForm()
        {
            InitializeComponent();            
        }

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

            this.Hide();
            try
            {
                string args = BuildArgs();
                string fileName = "BibleNoteLinker.exe";
                string filePath = Path.Combine(Utils.GetCurrentDirectory(), fileName);

                Process.Start(filePath, args);

                this.Close();
            }
            catch
            {
                this.Show();
                throw;
            }
        }
        
        private string BuildArgs()
        {
            StringBuilder sb = new StringBuilder();

            if (rbAnalyzeAllPages.Enabled && rbAnalyzeAllPages.Checked)
                sb.AppendFormat(" {0}", Arg_AllPages);
            else if (rbAnalyzeChangedPages.Enabled && rbAnalyzeChangedPages.Checked)
                sb.AppendFormat(" {0} {1}", Arg_AllPages, Arg_Changed);            

            if (chkForce.Enabled && chkForce.Checked)
                sb.AppendFormat(" {0}", Arg_Force);            

            return sb.ToString();
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
    }
}
