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
        private const double FormHeightProportion = 0.95;  // от всего экрана
        private const double FormWidthProportion = 0.33;

        private string _titleAtStart;
        private bool _suppressTbScaleLayout;
        private bool _touchInputAvailable;

        protected OpenBibleVerseHandler OpenBibleVerseHandler { get; set; }
        protected NavigateToHandler NavigateToHandler { get; set; }

        public bool ExitApplication { get; set; }   

        public NotesPageForm()
        {   
            this.SetFormUICulture();

            InitializeComponent();            

            OpenBibleVerseHandler = new OpenBibleVerseHandler();
            NavigateToHandler = new NavigateToHandler();

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

        public void OpenNotesPage(VersePointer vp, string verseNotesPageFilePath)
        {   
            if (!string.IsNullOrEmpty(verseNotesPageFilePath))
            {
                if (!File.Exists(verseNotesPageFilePath))
                    FormLogger.LogMessage(BibleCommon.Resources.Constants.VerseIsNotMentioned);
                else
                {
                    if (!vp.IsChapter && !SettingsManager.Instance.UseDifferentPagesForEachVerse)
                        verseNotesPageFilePath += "#" + vp.Verse.Value;

                    wbNotesPage.Url = new Uri(verseNotesPageFilePath);

                    if (!this.Visible)                    
                        this.Show();                                            

                    if (this.WindowState != FormWindowState.Normal)
                        this.WindowState = FormWindowState.Normal;

                    this.SetFocus();
                    wbNotesPage.Focus();

                    this.Text = string.Format("{0} ({1})", _titleAtStart, vp.GetFriendlyFullVerseName());
                }
            }
        }        

        private void NotesPageForm_Load(object sender, EventArgs e)
        {
            SetCheckboxes();
            SetLocation();
            SetSize();
            SetScale();            
            wbNotesPage.Focus();

            _touchInputAvailable = Utils.TouchInputAvailable();
        }        

        private void SetScale()
        {
            _suppressTbScaleLayout = true;
            tbScale.Text = Properties.Settings.Default.NotesPageFormScale.ToString();
            _suppressTbScaleLayout = false;
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
            Properties.Settings.Default.NotesPageFormScale = GetInputScale();

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
            tbScale_TextChanged(this, null);

            if (_touchInputAvailable)
            {
                var styleEl = wbNotesPage.Document.CreateElement("style");
                styleEl.SetAttribute("type", "text/css");
                styleEl.InnerHtml = " li.pageLevel { padding-bottom:5px; } .subLinks { padding-top:5px; } ";
                wbNotesPage.Document.Body.AppendChild(styleEl);
            }
        }

        private void tbScale_TextChanged(object sender, EventArgs e)
        {
            if (!_suppressTbScaleLayout)
            {
                var scale = GetInputScale();
                //var k = scale > 10 ? 0.1 : 0.05;
                //k = (float)(1 + (scale - 10) * k);
                wbNotesPage.Zoom(scale);

                if (!tbScale.Text.EndsWith("%"))
                {
                    _suppressTbScaleLayout = true;
                    tbScale.Text += "%";
                    _suppressTbScaleLayout = false;
                }
            }
        }

        private void tbScale_KeyPress(object sender, KeyPressEventArgs e)
        {   
            e.Handled = true;
        }

        private void btnScaleUp_Click(object sender, EventArgs e)
        {
            var scale = GetInputScale();
            if (scale < 200)
                tbScale.Text = (scale + 5).ToString();
        }

        private void btnScaleDown_Click(object sender, EventArgs e)
        {
            var scale = GetInputScale();
            if (scale > 50)
                tbScale.Text = (scale - 5).ToString();
        }

        private int GetInputScale()
        {
            var scale = tbScale.Text;
            if (scale.EndsWith("%"))
                scale = scale.Remove(scale.Length - 1);

            return Convert.ToInt32(scale);                
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
    }
}
