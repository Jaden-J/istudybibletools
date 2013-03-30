using System;
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

namespace ISBTCommandHandler
{
    public partial class NotesPageForm : Form
    {
        private const double FormHeightProportion = 0.95;
        private const double FormWidthProportion = 0.33;

        protected OpenBibleVerseHandler OpenBibleVerseHandler { get; set; }
        protected NavigateToHandler NavigateToHandler { get; set; }

        public bool ExitApplication { get; set; }

        public NotesPageForm()
        {            
            InitializeComponent();            

            OpenBibleVerseHandler = new OpenBibleVerseHandler();
            NavigateToHandler = new NavigateToHandler();
        }

        public void OpenNotesPage(string verseNotesPageFilePath)
        {   
            if (!string.IsNullOrEmpty(verseNotesPageFilePath))
            {
                if (!File.Exists(verseNotesPageFilePath))
                    FormLogger.LogMessage(BibleCommon.Resources.Constants.VerseIsNotMentioned);
                else
                {
                    wbNotesPage.Url = new Uri(verseNotesPageFilePath);

                    if (!this.Visible)
                        this.Show();                    

                    this.SetFocus();
                }
            }
        }        

        private void NotesPageForm_Load(object sender, EventArgs e)
        {
            SetCheckboxes();
            SetLocation();
            SetSize();            
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
                this.Location = new Point(x, y);
            }
            else
            {
                var screenInfo = Screen.FromControl(this).Bounds;
                this.Location = new Point(Convert.ToInt32(screenInfo.Size.Width * (1 - FormWidthProportion)), 0);
            }
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

            if (!ExitApplication)
                e.Cancel = true;
        }        
    }
}
