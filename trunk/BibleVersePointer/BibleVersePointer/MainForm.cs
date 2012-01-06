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

namespace BibleVersePointer
{
    public partial class MainForm : Form
    {   

        private Microsoft.Office.Interop.OneNote.Application _onenoteApp = null;
        public Microsoft.Office.Interop.OneNote.Application OneNoteApp
        {
            get
            {
                return _onenoteApp;
            }
        }

        public MainForm()
        {
            InitializeComponent();

            _onenoteApp = new Microsoft.Office.Interop.OneNote.Application();
        }

        [DllImport("user32.dll")]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        private void MainForm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                this.Close();
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            Logger.Initialize();


            if (!string.IsNullOrEmpty(tbVerse.Text))
            {
                btnOk.Enabled = false;
                System.Windows.Forms.Application.DoEvents();

                try
                {
                    VersePointer vp = new VersePointer(tbVerse.Text);

                    if (!vp.IsValid)
                        vp = VersePointer.GetChapterVersePointer(tbVerse.Text);

                    if (!vp.IsValid)
                        vp = new VersePointer(tbVerse.Text + " 1:0");  // может только название книги

                    if (vp.IsValid)
                    {
                        if (OneNoteApp.Windows.CurrentWindow == null)
                            OneNoteApp.NavigateTo(string.Empty);

                        if (GoToVerse(vp))
                        {
                            Properties.Settings.Default.LastVerse = tbVerse.Text;
                            Properties.Settings.Default.Save();
                        }
                    }
                    else
                        throw new Exception("Не удалось распознать строку");
                }
                catch (Exception ex)
                {
                    Logger.LogError(ex.Message);
                }

                btnOk.Enabled = true;
            }

            if (!Logger.WasLogged)
            {
                SetForegroundWindow(new IntPtr((long)OneNoteApp.Windows.CurrentWindow.WindowHandle));
                this.Close();
            }
        }

        private bool GoToVerse(VersePointer vp)
        {
            HierarchySearchManager.HierarchySearchResult result = HierarchySearchManager.GetHierarchyObject(
                OneNoteApp, OneNoteApp.Windows.CurrentWindow.CurrentNotebookId, vp);

            if (result.ResultType != HierarchySearchManager.HierarchySearchResultType.NotFound)
            {
                string hierarchyObjectId = !string.IsNullOrEmpty(result.HierarchyObjectInfo.PageId)
                    ? result.HierarchyObjectInfo.PageId : result.HierarchyObjectInfo.SectionId;

                OneNoteApp.NavigateTo(hierarchyObjectId, result.HierarchyObjectInfo.ContentObjectId);
                return true;
            }
            else
                Logger.LogError("Не удалось определить место");

            return false;
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            tbVerse.Text = (string)Properties.Settings.Default.LastVerse;               
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
