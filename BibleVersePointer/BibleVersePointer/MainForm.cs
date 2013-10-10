using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.OneNote;
using System.Xml.Linq;
using System.Xml;
using System.Xml.XPath;
using System.Collections.Specialized;
using System.Runtime.InteropServices;
using BibleCommon;
using System.Threading;
using BibleCommon.Common;
using BibleCommon.Services;
using BibleCommon.Helpers;
using BibleCommon.Consts;
using System.Diagnostics;
using BibleCommon.Handlers;

namespace BibleVersePointer
{
    public partial class MainForm : Form
    {   
        private Microsoft.Office.Interop.OneNote.Application _oneNoteApp;
        private bool _systemIsConfigured;
        private object _locker = new object();

        public MainForm()
        {
            this.SetFormUICulture();

            InitializeComponent();

            _oneNoteApp = OneNoteUtils.CreateOneNoteAppSafe();
            
            this.Text = BibleCommon.Resources.Constants.OpenVerse; 
            lblDescription.Text = BibleCommon.Resources.Constants.SpecifyBibleVerse;

            new Thread(InitializeWithLock).Start();
        }

        public void InitializeWithLock()
        {
            lock (_locker)
            {
                Initialize();
            }
        }

        public void Initialize()
        {
            _systemIsConfigured = SettingsManager.Instance.IsConfigured(ref _oneNoteApp); // разгоняем
        }

        private void MainForm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                this.Close();
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            FormLogger.Initialize();
            BibleCommon.Services.Logger.Init("BibleVersePointer");

            try
            {
                if (!_systemIsConfigured)
                {
                    // так как программа кэшируется в пуле OneNote, то проверим - может уже сконфигурили всё.
                    SettingsManager.Initialize();
                    ApplicationCache.Initialize();
                    
                    Initialize();

                    if (!_systemIsConfigured)
                        FormLogger.LogError(BibleCommon.Resources.Constants.Error_SystemIsNotConfigured);
                }
                else
                {
                    if (!string.IsNullOrEmpty(tbVerse.Text))
                    {
                        btnOk.Enabled = false;
                        System.Windows.Forms.Application.DoEvents();

                        try
                        {
                            VersePointer vp = new VersePointer(tbVerse.Text);

                            if (!vp.IsValid)
                                vp = new VersePointer(tbVerse.Text + " 1:0");  // может только название книги

                            if (vp.IsValid)
                            {
                                var url = OpenBibleVerseHandler.GetCommandUrlStatic(vp, null);
                                Process.Start(url);

                                //this.Visible = false;
                                Properties.Settings.Default.LastVerse = tbVerse.Text;
                                Properties.Settings.Default.Save();
                            }
                            else
                                throw new Exception(BibleCommon.Resources.Constants.BibleVersePointerCanNotParseString);
                        }
                        catch (Exception ex)
                        {
                            FormLogger.LogError(OneNoteUtils.ParseErrorAndMakeItMoreUserFriendly(ex.Message));
                            tbVerse.SelectAll();
                        }
                    }

                    btnOk.Enabled = true;
                }

                if (!FormLogger.WasErrorLogged)
                {
                    OneNoteUtils.SetActiveCurrentWindow(ref _oneNoteApp);
                  //  this.Close();
                }
            }
            finally
            {
                BibleCommon.Services.Logger.Done();
            }
        }     

        private void MainForm_Load(object sender, EventArgs e)
        {
            try
            {
                tbVerse.Text = (string)Properties.Settings.Default.LastVerse;
            }
            catch (Exception ex)
            {
                FormLogger.LogError(ex);
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

        private void MainForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            OneNoteUtils.ReleaseOneNoteApp(ref _oneNoteApp);
        }   
    }
}
