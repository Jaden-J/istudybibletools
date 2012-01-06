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

namespace BibleNoteLinkerEx
{
    public partial class MainForm : Form
    {
        const string Arg_AllPages = "-allpages";
        const string Arg_DeleteNotes = "-deletenotes";
        const string Arg_Force = "-force";

        public MainForm()
        {
            InitializeComponent();            

        }

        [DllImport("user32.dll")]
        static extern bool SetForegroundWindow(IntPtr hWnd);        

        private void btnOk_Click(object sender, EventArgs e)
        {
            if (!chkDeleteNotes.Checked)
            {
                Settings.Default.AllPages = rbAnalyzeAllPages.Checked;
                Settings.Default.Force = chkForce.Checked;
                Settings.Default.Save();
            }
            else if (MessageBox.Show("Удалить все сводные страницы заметок и ссылки на них?", "Внимание!", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) == System.Windows.Forms.DialogResult.No)
                return;            

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
            else if (chkDeleteNotes.Checked)
                sb.AppendFormat(" {0} {1}", Arg_AllPages, Arg_DeleteNotes);

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
            if (Settings.Default.AllPages)
                rbAnalyzeAllPages.Checked = true;

            if (Settings.Default.Force)
                chkForce.Checked = true;         
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

        private void cbDeleteNotes_CheckedChanged(object sender, EventArgs e)
        {
            rbAnalyzeAllPages.Enabled = 
                rbAnalyzeCurrentPage.Enabled = 
                    chkForce.Enabled = !chkDeleteNotes.Checked;
        }
    }
}
