﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using BibleCommon.Handlers;
using BibleCommon.Common;
using BibleCommon.Contracts;
using System.Runtime.InteropServices;
using BibleCommon.Services;
using System.IO;
using BibleCommon.Helpers;
using System.Diagnostics;

namespace ISBTCommandHandler
{
    [ComVisible(true)]
    public partial class NotesPageForm : Form
    {
        private const double FormHeightProportion = 0.95;  // от всего экрана
        private const double FormWidthProportion = 0.33;

        private string _titleAtStart;        
        private bool _touchInputAvailable;

        private string _verseNotesPageFilePath;
        protected string VerseNotesPageFilePath
        {
            get
            {
                return _verseNotesPageFilePath;
            }
            set
            {
                _verseNotesPageFilePath = value;
                _filesInCurrentDirectory = null;
            }
        }

        private OrderedDictionary<VerseNumber, string> _filesInCurrentDirectory;
        protected OrderedDictionary<VerseNumber, string> FilesInCurrentDirectory
        {
            get
            {
                if (_filesInCurrentDirectory == null)
                {
                    _filesInCurrentDirectory = new OrderedDictionary<VerseNumber, string>();

                    foreach(var file in Directory.GetFiles(Path.GetDirectoryName(VerseNotesPageFilePath)))
                    {
                        try
                        {
                            _filesInCurrentDirectory.Add(GetFileVerseNumber(file), file);
                        }
                        catch (Exception ex)
                        {
                            BibleCommon.Services.Logger.LogError(ex);
                        }
                    }

                    _filesInCurrentDirectory.SortKeys();
                }

                return _filesInCurrentDirectory;
            }
        }

        protected VersePointer VersePointer { get; set; }
        protected OpenBibleVerseHandler OpenBibleVerseHandler { get; set; }
        protected NavigateToHandler NavigateToHandler { get; set; }

        public bool ExitApplication { get; set; }   

        public NotesPageForm()
        {   
            this.SetFormUICulture();

            InitializeComponent();            

            OpenBibleVerseHandler = new OpenBibleVerseHandler();
            NavigateToHandler = new NavigateToHandler();
            wbNotesPage.ObjectForScripting = this;            

            _titleAtStart = this.Text;                        
        }

        protected override void OnMouseWheel(MouseEventArgs e)
        {
            try
            {
                base.OnMouseWheel(e);
                if (//(!wbNotesPage.Focused || !wbNotesPage.RectangleToScreen(wbNotesPage.ClientRectangle).Contains(Cursor.Position)) && 
                    this.RectangleToScreen(this.ClientRectangle).Contains(Cursor.Position))
                {
                    wbNotesPage.Scroll(e);
                }
            }
            catch (Exception ex)
            {
                FormLogger.LogError(ex);
            }
        }

        public void chkDetailedNodes_Changed(bool isChecked)
        {
            Properties.Settings.Default.NotesPageFormShowDetailedNotes = isChecked;
        }

        public void OpenNotesPage(VersePointer vp, string verseNotesPageFilePath)
        {
            try
            {
                this.VerseNotesPageFilePath = verseNotesPageFilePath;
                this.VersePointer = vp;

                if (!string.IsNullOrEmpty(VerseNotesPageFilePath))
                {
                    if (!File.Exists(VerseNotesPageFilePath))
                        FormLogger.LogMessage(BibleCommon.Resources.Constants.VerseIsNotMentioned);
                    else
                    {
                        //if (!vp.IsChapter && !SettingsManager.Instance.UseDifferentPagesForEachVerse)
                        //    verseNotesPageFilePath += "#" + vp.Verse.Value;

                        wbNotesPage.Url = new Uri(VerseNotesPageFilePath);

                        if (!this.Visible)
                            this.Show();

                        if (this.WindowState != FormWindowState.Normal)
                            this.WindowState = FormWindowState.Normal;

                        this.SetFocus();
                        wbNotesPage.Focus();

                        this.Text = string.Format("{0} ({1})", _titleAtStart, VersePointer.GetFriendlyFullVerseName());

                        SetNavigationButtonsAvailability();
                    }
                }
            }
            catch (Exception ex)
            {
                FormLogger.LogError(ex);
            }
        }        

        private void NotesPageForm_Load(object sender, EventArgs e)
        {
            SetCheckboxes();
            SetLocation();
            SetSize();                      
            wbNotesPage.Focus();

            _touchInputAvailable = Utils.TouchInputAvailable();
        }

        private void SetScale()
        {            
            wbNotesPage.Zoom(Properties.Settings.Default.NotesPageFormScale);
        }

        private void SetCheckboxes()
        {
            chkAlwaysOnTop.Checked = Properties.Settings.Default.NotesPageFormAlwaysOnTop;
            chkCloseOnClick.Checked = Properties.Settings.Default.NotesPageFormCloseOnClick;
        }

        private void SetSize()
        {
            var size = Properties.Settings.Default.NotesPageFormSize;
            if (!string.IsNullOrEmpty(size))
            {
                var sizeParts = size.Split(new char[] { ';' });
                var w = int.Parse(sizeParts[0]);
                var h = int.Parse(sizeParts[1]);
                this.Size = new Size(w, h);
            }
            else
            {
                var screenInfo = Screen.FromControl(this).Bounds;
                this.Size = new Size(
                                 Convert.ToInt32(screenInfo.Size.Width * FormWidthProportion), 
                                 Convert.ToInt32(screenInfo.Size.Height * FormHeightProportion));
            } 
        }

        private void SetLocation()
        {
            var position = Properties.Settings.Default.NotesPageFormPosition;
            if (!string.IsNullOrEmpty(position))
            {
                var positionParts = position.Split(new char[] { ';' });
                var x = int.Parse(positionParts[0]);
                var y = int.Parse(positionParts[1]);

                if (x >= 0 && y >= 0)                
                    this.Location = new Point(x, y);
                else
                    SetDefaultPosition();
            }
            else            
                SetDefaultPosition();            
        }

        private void SetDefaultPosition()
        {
            var screenInfo = Screen.FromControl(this).Bounds;
            this.Location = new Point(Convert.ToInt32(screenInfo.Size.Width * (1 - FormWidthProportion)), 0);
        }

        private void wbNotesPage_Navigating(object sender, WebBrowserNavigatingEventArgs e)
        {
            var url = e.Url.ToString();            

            if (url.StartsWith(BibleCommon.Consts.Constants.OneNoteProtocol, StringComparison.OrdinalIgnoreCase)
                || OpenBibleVerseHandler.IsProtocolCommand(url) || NavigateToHandler.IsProtocolCommand(url))
            {
                if (chkCloseOnClick.Checked)
                    this.Hide();                
            }                        
        }

        private void chkAlwaysOnTop_CheckedChanged(object sender, EventArgs e)
        {
            this.TopMost = chkAlwaysOnTop.Checked;
        }

        private void NotesPageForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            Properties.Settings.Default.NotesPageFormAlwaysOnTop = chkAlwaysOnTop.Checked;
            Properties.Settings.Default.NotesPageFormCloseOnClick = chkCloseOnClick.Checked;
            Properties.Settings.Default.NotesPageFormPosition = string.Format("{0};{1}", this.Left, this.Top);
            Properties.Settings.Default.NotesPageFormSize= string.Format("{0};{1}", this.Width, this.Height);            

            Properties.Settings.Default.Save();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Hide();
        }
        private void NotesPageForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Hide();

            if (e.CloseReason != CloseReason.WindowsShutDown && !ExitApplication)
                e.Cancel = true;
        }
        
        private void wbNotesPage_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            SetScale();            
            wbNotesPage.Document.InvokeScript("initDetailedNotes", new object[] { Properties.Settings.Default.NotesPageFormShowDetailedNotes });

            if (_touchInputAvailable)
            {
                var styleEl = wbNotesPage.Document.CreateElement("style");
                styleEl.SetAttribute("type", "text/css");
                styleEl.InnerHtml = " li.pageLevel { padding-bottom:5px; } .subLinks { padding-top:5px; } ";
                wbNotesPage.Document.Body.AppendChild(styleEl);
            }
        }                        

        private void btnScaleUp_Click(object sender, EventArgs e)
        {
            if (Properties.Settings.Default.NotesPageFormScale < 200)
                Properties.Settings.Default.NotesPageFormScale += 5;

            SetScale();
        }

        private void btnScaleDown_Click(object sender, EventArgs e)
        {
            if (Properties.Settings.Default.NotesPageFormScale > 50)
                Properties.Settings.Default.NotesPageFormScale -= 5;

            SetScale();
        }        

        private bool _firstShown = true;
        private void NotesPageForm_Shown(object sender, EventArgs e)
        {
            if (_firstShown)
            {
                this.Focus();                
                _firstShown = false;
            }            
        }
        
        private void SetNavigationButtonsAvailability()
        {
            if (FilesInCurrentDirectory.First().Value != VerseNotesPageFilePath)
                btnPrev.Enabled = true;
            else
                btnPrev.Enabled = false;

            if (FilesInCurrentDirectory.Last().Value != VerseNotesPageFilePath)
                btnNext.Enabled = true;
            else
                btnNext.Enabled = false;
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            var nextFile = FilesInCurrentDirectory[FilesInCurrentDirectory.IndexOf(GetFileVerseNumber(VerseNotesPageFilePath)) + 1];
            OpenAnotherFile(nextFile);
        }

        private void btnPrev_Click(object sender, EventArgs e)
        {
            var prevFile = FilesInCurrentDirectory[FilesInCurrentDirectory.IndexOf(GetFileVerseNumber(VerseNotesPageFilePath)) - 1];
            OpenAnotherFile(prevFile);
        }

        private void OpenAnotherFile(string file)
        {
            var verseNumber = VerseNumber.Parse(Path.GetFileNameWithoutExtension(file));
            var vp = new VersePointer(VersePointer, verseNumber);
            OpenNotesPage(vp, file);            
        }

        private static VerseNumber GetFileVerseNumber(string file)
        {
            return VerseNumber.Parse(Path.GetFileNameWithoutExtension(file));
        }
    }
}