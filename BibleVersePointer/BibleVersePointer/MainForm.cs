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
        private Microsoft.Office.Interop.OneNote.Application _oneNoteApp = null;       

        public MainForm()
        {
            this.SetFormUICulture();

            InitializeComponent();

            _oneNoteApp = new Microsoft.Office.Interop.OneNote.Application();
            
            this.Text = BibleCommon.Resources.Constants.OpenVerse; 
            lblDescription.Text = BibleCommon.Resources.Constants.SpecifyBibleVerse;
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
            BibleCommon.Services.Logger.Init("BibleVersePointer");

            try
            {
                //if (!SettingsManager.Instance.IsConfigured(ref _oneNoteApp))
                //{
                //    SettingsManager.Instance.ReLoadSettings();  // так как программа кэшируется в пуле OneNote, то проверим - может уже сконфигурили всё.

                //    if (!SettingsManager.Instance.IsConfigured(ref _oneNoteApp))
                //    {
                //        Logger.LogError(BibleCommon.Resources.Constants.Error_SystemIsNotConfigured);
                //    }
                //}
                //else
                //{
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
                                var handler = new OpenBibleVerseHandler();
                                var url = handler.GetCommandUrl(vp, null);
                                Process.Start(url);


                                //OneNoteUtils.UseOneNoteAPI(ref _oneNoteApp, () =>
                                //{
                                //    if (_oneNoteApp.Windows.CurrentWindow == null)
                                //        _oneNoteApp.NavigateTo(string.Empty);
                                //});

                                //if (GoToVerse(vp))
                                //{
                                    this.Visible = false;
                                    Properties.Settings.Default.LastVerse = tbVerse.Text;
                                    Properties.Settings.Default.Save();
                                //}
                            }
                            else
                                throw new Exception(BibleCommon.Resources.Constants.BibleVersePointerCanNotParseString);
                        }
                        catch (Exception ex)
                        {
                            Logger.LogError(OneNoteUtils.ParseError(ex.Message));
                        }
                    }

                    btnOk.Enabled = true;
                //}

                if (!Logger.WasLogged)
                {
                    OneNoteUtils.UseOneNoteAPI(ref _oneNoteApp, () =>
                    {
                        if (_oneNoteApp.Windows.CurrentWindow != null)
                            SetForegroundWindow(new IntPtr((long)_oneNoteApp.Windows.CurrentWindow.WindowHandle));
                    });
                    this.Close();
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
            _oneNoteApp = null;
        }

        private bool GoToVerse(VersePointer vp)
        {
            var result = HierarchySearchManager.GetHierarchyObject(ref _oneNoteApp, SettingsManager.Instance.NotebookId_Bible, vp, HierarchySearchManager.FindVerseLevel.OnlyFirstVerse);

            if (result.ResultType != HierarchySearchManager.HierarchySearchResultType.NotFound
                && (result.HierarchyStage == HierarchySearchManager.HierarchyStage.ContentPlaceholder || result.HierarchyStage == HierarchySearchManager.HierarchyStage.Page))
            {
                string hierarchyObjectId = !string.IsNullOrEmpty(result.HierarchyObjectInfo.PageId)
                    ? result.HierarchyObjectInfo.PageId : result.HierarchyObjectInfo.SectionId;

                NavigateTo(hierarchyObjectId, result.HierarchyObjectInfo.GetAllObjectsIds().ToArray());
                return true;
            }
            else
                Logger.LogError(BibleCommon.Resources.Constants.BibleVersePointerCanNotFindPlace);

            return false;
        }

        private void NavigateTo(string pageId, params HierarchySearchManager.VerseObjectInfo[] objectsIds)
        {
            if (objectsIds.Length > 0 && !string.IsNullOrEmpty(objectsIds[0].ObjectHref))
            {
                Process.Start(objectsIds[0].ObjectHref);   // иначе, если делать через NavigateTo, то когда, например, дропбокс изменит имя файла секции (сделает маленькими буквами) - ID меняется и выдаётся ошибка.
            }
            else
            {
                OneNoteUtils.UseOneNoteAPI(ref _oneNoteApp, () =>
                {
                    _oneNoteApp.NavigateTo(pageId, objectsIds.Length > 0 ? objectsIds[0].ObjectId : null);
                });
            }

            if (objectsIds.Length > 1)
            {   
                XmlNamespaceManager xnm;                
                var pageDoc = OneNoteUtils.GetPageContent(ref _oneNoteApp, pageId, PageInfo.piSelection, out xnm);
                OneNoteLocker.UnlockCurrentSection(ref _oneNoteApp);
                
                foreach (var objectId in objectsIds.Skip(1))
                {
                    var el = pageDoc.Root.XPathSelectElement(string.Format("//one:OE[@objectID=\"{0}\"]/one:T", objectId), xnm);
                    if (el != null)
                        el.SetAttributeValue("selected", "all");
                }
                
                OneNoteUtils.UpdatePageContentSafe(ref _oneNoteApp, pageDoc, xnm);
            }
        }      
    }
}
