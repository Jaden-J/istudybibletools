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

namespace BibleVerseLinkerEx
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        [DllImport("user32.dll")]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        private void btnOk_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.LastPageName = tbPageName.Text;
            Properties.Settings.Default.LastSearchForUnderlineText = cbSearchForUnderlineText.Checked;
            Properties.Settings.Default.Save();

            btnOk.Enabled = false;
            Application.DoEvents();

            Logger.Initialize();

            if (!SettingsManager.Instance.IsConfigured())
            {
                Logger.LogError("Система не сконфигурирована");
            }
            else
            {
                try
                {
                    VerseLinker vlManager = new VerseLinker();
                    vlManager.SearchForUnderlineText = cbSearchForUnderlineText.Checked;
                    if (!string.IsNullOrEmpty(tbPageName.Text))
                        vlManager.DescriptionPageName = tbPageName.Text;

                    vlManager.Do();

                    if (!Logger.WasLogged)
                        SetForegroundWindow(new IntPtr((long)vlManager.OneNoteApp.Windows.CurrentWindow.WindowHandle));
                }
                catch (Exception ex)
                {
                    Logger.LogError(ex.Message);
                }
            }

            btnOk.Enabled = true;

            if (!Logger.WasLogged)                
                this.Close();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            tbPageName.Text = Properties.Settings.Default.LastPageName;
            cbSearchForUnderlineText.Checked = Properties.Settings.Default.LastSearchForUnderlineText;           
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
