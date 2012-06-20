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

        [DllImport("user32.dll")]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        private void btnOk_Click(object sender, EventArgs e)
        {
            btnOk.Enabled = false;
            Application.DoEvents();

            Microsoft.Office.Interop.OneNote.Application oneNoteApp = new Microsoft.Office.Interop.OneNote.Application();

            Logger.Initialize();

            if (!SettingsManager.Instance.IsConfigured(oneNoteApp))
            {
                Logger.LogError(BibleCommon.Resources.Constants.Error_SystemIsNotConfigures);
            }
            else
            {
                try
                {
                    BibleCommon.Services.OneNoteLocker.UnlockCurrentSection(oneNoteApp);
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

                        if (!Logger.WasLogged)
                        {
                            SetForegroundWindow(new IntPtr((long)vlManager.OneNoteApp.Windows.CurrentWindow.WindowHandle));
                            this.Visible = false;
                            vlManager.SortCommentsPages();
                        }
                    }
                }
                catch (Exception ex)
                {
                    Logger.LogError(ex.Message);
                }
            }

            btnOk.Enabled = true;

            if (!Logger.WasLogged)
            {
                this.Visible = false;
                Properties.Settings.Default.LastPageName = tbPageName.Text;                
                Properties.Settings.Default.Save();
                this.Close();
            }
        }
       

        private void MainForm_Load(object sender, EventArgs e)
        {
            tbPageName.Text = Properties.Settings.Default.LastPageName;            
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
