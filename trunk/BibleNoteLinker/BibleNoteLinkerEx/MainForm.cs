using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using BibleNoteLinkerEx.Properties;
using System.Configuration;
using BibleCommon;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using BibleCommon.Helpers;
using BibleCommon.Services;
using System.Xml.Linq;
using BibleCommon.Consts;

namespace BibleNoteLinkerEx
{
    public partial class MainForm : Form
    {
        private Microsoft.Office.Interop.OneNote.Application _oneNoteApp;
 
        const string Arg_AllPages = "-allpages";        
        const string Arg_Changed = "-changed";
        const string Arg_Force = "-force";                

        public MainForm()
        {
            InitializeComponent();
            _oneNoteApp = new Microsoft.Office.Interop.OneNote.Application();
        }


        private int _originalFormHeight;
        const int FirstFormHeight = 185;
        const int SecondFormHeight = 250;

        private delegate void SetControlPropertyThreadSafeDelegate(Control control, string propertyName, object propertyValue);

        public static void SetControlPropertyThreadSafe(Control control, string propertyName, object propertyValue)
        {
            if (control.InvokeRequired)
            {
                control.Invoke(new SetControlPropertyThreadSafeDelegate(SetControlPropertyThreadSafe), new object[] { control, propertyName, propertyValue });
            }
            else
            {
                control.GetType().InvokeMember(propertyName, BindingFlags.SetProperty, null, control, new object[] { propertyValue });
            }
        }

        [DllImport("user32.dll")]
        static extern bool SetForegroundWindow(IntPtr hWnd);


        private void btnOk_Click(object sender, EventArgs e)
        {
            BibleNoteLinkerEx.Properties.Settings.Default.AllPages = rbAnalyzeAllPages.Checked;
            BibleNoteLinkerEx.Properties.Settings.Default.Changed = rbAnalyzeChangedPages.Checked;
            BibleNoteLinkerEx.Properties.Settings.Default.Force = chkForce.Checked;
            BibleNoteLinkerEx.Properties.Settings.Default.Save();

            this.Height = SecondFormHeight;
            EnableBaseElements(false);

            BibleCommon.Services.Logger.Init("BibleNoteLinkerEx");
            //BibleCommon.Services.Logger.SetOutputListBox(lbLog);

            try
            {
                if (!rbAnalyzeCurrentPage.Checked)
                {
                    List<NotebookIterator.NotebookInfo> notebooks = GetNotebooksInfo();
                    pbMain.Maximum = notebooks.Sum(notebook => notebook.PagesCount);
                    foreach (NotebookIterator.NotebookInfo notebook in notebooks)
                        ProcessNotebook(notebook);
                }
                else
                {
                    pbMain.Maximum = 1;

                }                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                EnableBaseElements(true);
                BibleCommon.Services.Logger.Done();
            }
        }
        private void LogMessage(string message)
        {
            lblProgress.Text = message;
            lbLog.Items.Add(message);
        }

        public void ProcessNotebook(NotebookIterator.NotebookInfo notebook)
        {   

            BibleCommon.Services.Logger.LogMessage("Обработка записной книжки: '{0}'", notebook.Title);

            BibleCommon.Services.Logger.MoveLevel(1);
            ProcessSectionGroup(notebook.RootSectionGroup, true);
            BibleCommon.Services.Logger.MoveLevel(-1);
        }

        private void ProcessSectionGroup(BibleCommon.Services.NotebookIterator.SectionGroupInfo sectionGroup, bool isRoot)
        {
            if (!isRoot)
            {
                BibleCommon.Services.Logger.LogMessage("Обработка группы секций '{0}'", sectionGroup.Title);
                BibleCommon.Services.Logger.MoveLevel(1);
            }

            foreach (BibleCommon.Services.NotebookIterator.SectionInfo section in sectionGroup.Sections)
            {
                BibleCommon.Services.Logger.LogMessage("Обработка секции '{0}'", section.Title);
                BibleCommon.Services.Logger.MoveLevel(1);

                foreach (BibleCommon.Services.NotebookIterator.PageInfo page in section.Pages)
                {
                    BibleCommon.Services.Logger.LogMessage("Обработка страницы '{0}'", page.Title);
                    BibleCommon.Services.Logger.MoveLevel(1);

                    NoteLinkManager noteLinkManager = new NoteLinkManager(_oneNoteApp);
                    noteLinkManager.LinkPageVerses(page.SectionGroupId, page.SectionId, page.Id, NoteLinkManager.AnalyzeDepth.Full, chkForce.Checked);
                    pbMain.PerformStep();
                    System.Windows.Forms.Application.DoEvents();

                    BibleCommon.Services.Logger.MoveLevel(-1);
                }

                BibleCommon.Services.Logger.MoveLevel(-1);
            }

            foreach (BibleCommon.Services.NotebookIterator.SectionGroupInfo subSectionGroup in sectionGroup.SectionGroups)
            {
                ProcessSectionGroup(subSectionGroup, false);
            }

            if (!isRoot)
                BibleCommon.Services.Logger.MoveLevel(-1);
        }


        private List<NotebookIterator.NotebookInfo> GetNotebooksInfo()
        {
            NotebookIterator iterator = new NotebookIterator(_oneNoteApp);
            List<NotebookIterator.NotebookInfo> result = new List<NotebookIterator.NotebookInfo>();

            Func<NotebookIterator.PageInfo, bool> filter = null;
            if (rbAnalyzeChangedPages.Checked)
                filter = IsPageWasModifiedAfterLastAnalyze;

            foreach (string id in Helper.GetSelectedNotebooksIds())
            {
                if (SettingsManager.Instance.IsSingleNotebook)                
                    result.Add(iterator.GetNotebookPages(SettingsManager.Instance.NotebookId_Bible, id, filter));                
                else                
                    result.Add(iterator.GetNotebookPages(id, null, filter));                
            }

            return result;
        }

        private bool IsPageWasModifiedAfterLastAnalyze(NotebookIterator.PageInfo page)
        {   
            XAttribute lastModifiedDateAttribute = page.PageElement.Attribute("lastModifiedTime");
            if (lastModifiedDateAttribute != null)
            {
                DateTime lastModifiedDate = DateTime.Parse(lastModifiedDateAttribute.Value);

                string lastAnalyzeTime = OneNoteUtils.GetPageMetaData(_oneNoteApp, page.PageElement, Constants.Key_LatestAnalyzeTime, page.Xnm);
                if (!string.IsNullOrEmpty(lastAnalyzeTime) && lastModifiedDate <= DateTime.Parse(lastAnalyzeTime).ToLocalTime())
                    return false;                
            }

            return true;
        }

        private void EnableBaseElements(bool enabled)
        {
            btnOk.Enabled = enabled;
            rbAnalyzeAllPages.Enabled = enabled;
            rbAnalyzeChangedPages.Enabled = enabled;
            rbAnalyzeCurrentPage.Enabled = enabled;
            chkForce.Enabled = enabled;
            tsmiSeelctNotebooks.Enabled = enabled;
        }
        
        private string BuildArgs()
        {
            StringBuilder sb = new StringBuilder();

            if (rbAnalyzeAllPages.Enabled && rbAnalyzeAllPages.Checked)
                sb.AppendFormat(" {0}", Arg_AllPages);
            else if (rbAnalyzeChangedPages.Enabled && rbAnalyzeChangedPages.Checked)
                sb.AppendFormat(" {0} {1}", Arg_AllPages, Arg_Changed);            

            if (chkForce.Enabled && chkForce.Checked)
                sb.AppendFormat(" {0}", Arg_Force);            

            return sb.ToString();
        }

        private void MainForm_KeyDown(object sender, KeyEventArgs e)
        {
             switch (e.KeyCode)
            {
                case Keys.Escape:
                    this.Close();
                    break;               
                case Keys.Space:
                    if (chkForce.Enabled)
                        chkForce.Checked = !chkForce.Checked;
                    e.SuppressKeyPress = true;
                    break;
            }
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            if (BibleNoteLinkerEx.Properties.Settings.Default.AllPages)
                rbAnalyzeAllPages.Checked = true;
            else if (BibleNoteLinkerEx.Properties.Settings.Default.Changed)
                rbAnalyzeChangedPages.Checked = true;

            if (BibleNoteLinkerEx.Properties.Settings.Default.Force)
                chkForce.Checked = true;

            lblInfo.Visible = false;
            _originalFormHeight = this.Height;
            this.Height = FirstFormHeight;

            new Thread(CheckForNewerVersion).Start();
        }       

        public void CheckForNewerVersion()
        {
            if (VersionOnServerManager.NeedToUpdate())
            {
                SetControlPropertyThreadSafe(lblInfo, "Text",
@"Доступна новая версия программы
на сайте http://IStudyBibleTools.ru. 
Кликните, чтобы перейти на страницу загрузки.");

                SetControlPropertyThreadSafe(this, "Size", new Size(this.Size.Width, this.Size.Height + 50));
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

        private void lblInfo_Click(object sender, EventArgs e)
        {
            Process.Start(BibleCommon.Consts.Constants.DownloadPageUrl);
        }

        private bool _detailsWereShown = false;
        private void llblDetails_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (!_detailsWereShown)
            {
                llblDetails.Text = "Скрыть детали";
                this.Height = _originalFormHeight;                
            }
            else
            {
                llblDetails.Text = "Показать детали";
                this.Height = SecondFormHeight;                
            }

            _detailsWereShown = !_detailsWereShown;
        }

        private void tsmiSeelctNotebooks_Click(object sender, EventArgs e)
        {
            SelectNoteBooks form = new SelectNoteBooks(_oneNoteApp);
            form.ShowDialog();
        }
    }
}
