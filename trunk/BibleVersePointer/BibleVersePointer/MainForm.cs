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
            this.SetFormUICulture();

            InitializeComponent();

            _onenoteApp = new Microsoft.Office.Interop.OneNote.Application();


            //todo: переделать на ресурсы. сейчас есть проблема с загрузкой сателлитных сборок, когда загружаем порграмму в пул OneNote
            this.Text = SettingsManager.Instance.Language == 1049 ? "Открыть стих" : "Open a verse";
            lblDescription.Text = SettingsManager.Instance.Language == 1049 ? "Укажите место Писания" : "Specify the Bible verse";
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

            if (!SettingsManager.Instance.IsConfigured(OneNoteApp))
            {
                SettingsManager.Instance.ReLoadSettings();  // так как программа кэшируется в пуле OneNote, то проверим - может уже сконфигурили всё.

                if (!SettingsManager.Instance.IsConfigured(OneNoteApp))
                {
                    Logger.LogError(BibleCommon.Resources.Constants.Error_SystemIsNotConfigures);
                }
                
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
                            if (OneNoteApp.Windows.CurrentWindow == null)
                                OneNoteApp.NavigateTo(string.Empty);

                            if (GoToVerse(vp))
                            {
                                this.Visible = false;
                                Properties.Settings.Default.LastVerse = tbVerse.Text;
                                Properties.Settings.Default.Save();
                            }
                        }
                        else
                            throw new Exception(BibleCommon.Resources.Constants.BibleVersePointerCanNotParseString);
                    }
                    catch (Exception ex)
                    {
                        Logger.LogError(ex.Message);
                    }
                }

                btnOk.Enabled = true;
            }

            if (!Logger.WasLogged)
            {
                if (OneNoteApp.Windows.CurrentWindow != null)                
                    SetForegroundWindow(new IntPtr((long)OneNoteApp.Windows.CurrentWindow.WindowHandle));
                this.Close();
            }
        }

        private bool GoToVerse(VersePointer vp)
        {            
            HierarchySearchManager.HierarchySearchResult result = HierarchySearchManager.GetHierarchyObject(OneNoteApp, SettingsManager.Instance.NotebookId_Bible, vp, false);

            if (result.ResultType != HierarchySearchManager.HierarchySearchResultType.NotFound)
            {
                string hierarchyObjectId = !string.IsNullOrEmpty(result.HierarchyObjectInfo.PageId)
                    ? result.HierarchyObjectInfo.PageId : result.HierarchyObjectInfo.SectionId;

                NavigateTo(OneNoteApp, hierarchyObjectId, result.HierarchyObjectInfo.GetAllObjectsIds().ToArray());
                return true;
            }
            else
                Logger.LogError(BibleCommon.Resources.Constants.BibleVersePointerCanNotFindPlace);

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

        private void MainForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            _onenoteApp = null;
        }


        private static void NavigateTo(Microsoft.Office.Interop.OneNote.Application oneNoteApp, string pageId, params string[] objectsIds)
        {
            oneNoteApp.NavigateTo(pageId, objectsIds.Length > 0 ? objectsIds[0] : null);            

            if (objectsIds.Length > 1)
            {   
                XmlNamespaceManager xnm;                
                var pageDoc = OneNoteUtils.GetPageContent(oneNoteApp, pageId, PageInfo.piSelection, out xnm);
                OneNoteLocker.UnlockCurrentSection(oneNoteApp);
                
                foreach (var objectId in objectsIds.Skip(1))
                {
                    var el = pageDoc.Root.XPathSelectElement(string.Format("//one:OE[@objectID='{0}']/one:T", objectId), xnm);
                    if (el != null)
                        el.SetAttributeValue("selected", "all");
                }
                
                OneNoteUtils.UpdatePageContentSafe(oneNoteApp, pageDoc, xnm);
            }
        }
    }
}
