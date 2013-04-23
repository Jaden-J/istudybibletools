using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using BibleCommon.Services;
using Microsoft.Office.Interop.OneNote;
using System.Xml.Linq;
using System.Xml.XPath;
using BibleCommon.Helpers;

namespace BibleNoteLinker
{
    public partial class SelectNoteBooksForm : Form
    {
        private Microsoft.Office.Interop.OneNote.Application _oneNoteApp;

        public SelectNoteBooksForm()
        {
            InitializeComponent();
            _oneNoteApp = OneNoteUtils.CreateOneNoteAppSafe();
        }

        private void SelectNoteBooks_Load(object sender, EventArgs e)
        {
            try
            {
                Dictionary<string, string> allNotebooks = GetAllNotebooks();                

                int i = 0;
                foreach (string id in allNotebooks.Keys)
                {
                    RenderSelectRow(id, allNotebooks[id],  i++);
                }

                SetElementsAttributes(i);

                LoadSelectedNotebooks();
            }
            catch (Exception ex)
            {
                FormLogger.LogError(ex);
            }   
        }

        private void LoadSelectedNotebooks()
        {
            foreach (Control chk in pnMain.Controls)
            {
                if (chk is CheckBox)
                {
                    if (SettingsManager.Instance.SelectedNotebooksForAnalyze.Exists(notebook => notebook.NotebookId == chk.Name))                                        
                    {
                        var notebookInfo = SettingsManager.Instance.SelectedNotebooksForAnalyze.FirstOrDefault(notebook => notebook.NotebookId == chk.Name);
                        ((CheckBox)chk).Checked = true;
                        if (notebookInfo.DisplayLevels.HasValue)
                        {
                            var cb = ((ComboBox)pnMain.Controls[GetDisplayLevelsControlId(chk.Name)]);
                            cb.SelectedItem = notebookInfo.DisplayLevels.ToString();
                        }
                    }
                    else
                        pnMain.Controls[GetDisplayLevelsControlId(chk.Name)].Visible = false;
                }
            }
        }

        private string GetDisplayLevelsControlId(string notebookId)
        {
            return notebookId + "_";
        }

        private void SetElementsAttributes(int notebooksCount)
        {
            int notebooksRowHeight = 25 * notebooksCount;
            btnOk.Top = notebooksRowHeight + 20 + 13;
            this.Height = btnOk.Top + btnOk.Height + 35;
            pnMain.Height = notebooksRowHeight;
            lblError.Top = notebooksRowHeight + 25;
        }

        private List<string> GetDisplayLevels()
        {
            var displayLevels = new List<string>() { BibleCommon.Resources.Constants.All };
            for (int level = 1; level <= 10; level++)
                displayLevels.Add(level.ToString());            

            return displayLevels;
        }

        private void RenderSelectRow(string id, string title, int index)
        {
            CheckBox chk = new CheckBox();
            chk.Name = id;
            chk.Text = title;
            chk.Font = new System.Drawing.Font(BibleCommon.Consts.Constants.UnicodeFontName, chk.Font.Size, chk.Font.Style);
            chk.Top = 25 * index;
            chk.Width = 330;
            chk.CheckedChanged += chk_CheckedChanged;
            pnMain.Controls.Add(chk);

            ComboBox cb = new ComboBox();
            cb.Name = GetDisplayLevelsControlId(id);
            cb.DataSource = GetDisplayLevels();
            cb.Top = 25 * index;
            cb.Left = 335;
            cb.Width = 45;
            pnMain.Controls.Add(cb);
        }

        void chk_CheckedChanged(object sender, EventArgs e)
        {
            var chk = (CheckBox)sender;
            
            pnMain.Controls[GetDisplayLevelsControlId(chk.Name)].Visible = chk.Checked;
        }

        private Dictionary<string, string> GetAllNotebooks()
        {
            Dictionary<string, string> result = new Dictionary<string, string>();

            if (SettingsManager.Instance.IsSingleNotebook)
            {
                ApplicationCache.HierarchyElement sectionGroups = ApplicationCache.Instance.GetHierarchy(ref _oneNoteApp, 
                    SettingsManager.Instance.NotebookId_Bible, HierarchyScope.hsChildren);

                foreach (XElement sectionGroup in sectionGroups.Content.Root.XPathSelectElements(string.Format("one:SectionGroup[{0}]", OneNoteUtils.NotInRecycleXPathCondition), sectionGroups.Xnm))
                {
                    result.Add((string)sectionGroup.Attribute("ID"), (string)sectionGroup.Attribute("name"));
                }
            }
            else
            {
                result = OneNoteUtils.GetExistingNotebooks(ref _oneNoteApp);                
            }

            return result;
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            var selectedNotebooks = GetSelectedNotebooks();
            if (selectedNotebooks.Count == 0)
            {
                lblError.Text = BibleCommon.Resources.Constants.NoteLinkerNoElementSelected;
                lblError.Visible = true;                
            }
            else
            {
                SettingsManager.Instance.SelectedNotebooksForAnalyze = selectedNotebooks;
                SettingsManager.Instance.Save();
                this.DialogResult = System.Windows.Forms.DialogResult.OK;
                this.Close();
            }
        }

        private List<NotebookForAnalyzeInfo> GetSelectedNotebooks()
        {
            var result = new List<NotebookForAnalyzeInfo>();

            foreach (Control chk in pnMain.Controls)
            {
                if (chk is CheckBox)
                {
                    if (((CheckBox)chk).Checked)
                    {
                        var notebookInfo = new NotebookForAnalyzeInfo(chk.Name);

                        var displayLevelsString = (string)((ComboBox)pnMain.Controls[GetDisplayLevelsControlId(chk.Name)]).SelectedItem;
                        int displayLevels;
                        if (int.TryParse(displayLevelsString, out displayLevels))
                            notebookInfo.DisplayLevels = displayLevels;

                        result.Add(notebookInfo);
                    }
                }
            }

            return result;
        }

        private void SelectNoteBooksForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            OneNoteUtils.ReleaseOneNoteApp(ref _oneNoteApp);
        }
    }
}
