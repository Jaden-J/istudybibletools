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

namespace ISBTCommandHandler
{
    public partial class NotesPageForm : Form
    {
        public NotesPageForm()
        {
            InitializeComponent();
        }

        public void OpenNotesPage(VersePointer vp)
        {
            var verseNotesPageFilePath = GetVerseNotesPageFilePath(vp);

            if (!string.IsNullOrEmpty(verseNotesPageFilePath))
            {
                if (!this.Visible)
                    this.Show();
                else
                    this.Focus();

                wbNotesPage.Url = new Uri(verseNotesPageFilePath);
            }
        }

        private string GetVerseNotesPageFilePath(VersePointer vp)
        {
            return @"C:\Users\lux_demko\Desktop\temp\temp\test.htm";
        }

        private void NotesPageForm_Load(object sender, EventArgs e)
        {
            
        }

        private void wbNotesPage_Navigating(object sender, WebBrowserNavigatingEventArgs e)
        {
            var url = e.Url.ToString();

            if (url.StartsWith(BibleCommon.Consts.Constants.OneNoteProtocol, StringComparison.OrdinalIgnoreCase)
                || url.StartsWith(NavigateToHandler.ProtocolFullString, StringComparison.OrdinalIgnoreCase))
            {
                if (chkCloseOnClick.Checked)
                    this.Hide();
            }
        }

        private void chkAlwaysOnTop_CheckedChanged(object sender, EventArgs e)
        {
            this.TopMost = chkAlwaysOnTop.Checked;
        }           
    }
}
