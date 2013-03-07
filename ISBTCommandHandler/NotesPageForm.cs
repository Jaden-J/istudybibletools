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

namespace ISBTCommandHandler
{
    public partial class NotesPageForm : Form
    {
        protected OpenBibleVerseHandler OpenBibleVerseHandler { get; set; }
        protected NavigateToHandler NavigateToHandler { get; set; }        
        public NotesPageForm()
        {            
            InitializeComponent();            

            OpenBibleVerseHandler = new OpenBibleVerseHandler();
            NavigateToHandler = new NavigateToHandler();
        }

        public void OpenNotesPage(VersePointer vp)
        {
            var verseNotesPageFilePath = GetVerseNotesPageFilePath(vp);

            if (!string.IsNullOrEmpty(verseNotesPageFilePath))
            {
                if (!File.Exists(verseNotesPageFilePath))
                    FormLogger.LogMessage(BibleCommon.Resources.Constants.VerseIsNotMentioned);
                else
                {
                    wbNotesPage.Url = new Uri(verseNotesPageFilePath);

                    if (!this.Visible)
                        this.Show();
                    else
                        this.Focus();
                }
            }
        }

        private string GetVerseNotesPageFilePath(VersePointer vp)
        {
            return OpenNotesPageHandler.GetNotesPageFilePath(vp, 
                SettingsManager.Instance.UseDifferentPagesForEachVerse ? NoteLinkManager.NotesPageType.Verse : NoteLinkManager.NotesPageType.Chapter);
        }

        private void NotesPageForm_Load(object sender, EventArgs e)
        {
            
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
            
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Hide();
        }
    }
}
