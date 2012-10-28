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

namespace BibleNoteLinker
{
    public partial class SelectNoteBooksForm : Form
    {
        private Microsoft.Office.Interop.OneNote.Application _oneNoteApp;

        public SelectNoteBooksForm(Microsoft.Office.Interop.OneNote.Application oneNoteApp)
        {
            InitializeComponent();
            _oneNoteApp = oneNoteApp;
        }

        private void SelectNoteBooks_Load(object sender, EventArgs e)
        {
            try
            {
                Dictionary<string, string> allNotebooks = GetAllNotebooks();

                int i = 0;
                foreach (string id in allNotebooks.Keys)
                {
                    RenderSelectRow(id, allNotebooks[id], i++);
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
            List<string> notebooksIds = Helper.GetSelectedNotebooksIds();

            foreach (CheckBox chk in pnMain.Controls)
            {
                if (notebooksIds.Contains(chk.Name))                
                    chk.Checked = true;                
            }
        }

        private void SetElementsAttributes(int notebooksCount)
        {
            int notebooksRowHeight = 25 * notebooksCount;
            btnOk.Top = notebooksRowHeight + 20;
            this.Height = btnOk.Top + btnOk.Height + 35;
            pnMain.Height = notebooksRowHeight;
            lblError.Top = notebooksRowHeight + 25;
        }

        private void RenderSelectRow(string id, string title, int index)
        {
            CheckBox cb = new CheckBox();
            cb.Name = id;
            cb.Text = title;
            cb.Top = 25 * index;
            cb.Width = 250;
            pnMain.Controls.Add(cb);
        }

        private Dictionary<string, string> GetAllNotebooks()
        {
            Dictionary<string, string> result = new Dictionary<string, string>();

            if (SettingsManager.Instance.IsSingleNotebook)
            {
                OneNoteProxy.HierarchyElement sectionGroups = OneNoteProxy.Instance.GetHierarchy(_oneNoteApp, 
                    SettingsManager.Instance.NotebookId_Bible, HierarchyScope.hsChildren);

                foreach (XElement sectionGroup in sectionGroups.Content.Root.XPathSelectElements("one:SectionGroup", sectionGroups.Xnm))
                {
                    result.Add((string)sectionGroup.Attribute("ID"), (string)sectionGroup.Attribute("name"));
                }
            }
            else
            {
                OneNoteProxy.HierarchyElement notebooks = OneNoteProxy.Instance.GetHierarchy(_oneNoteApp, null, HierarchyScope.hsNotebooks);

                foreach (XElement notebook in notebooks.Content.Root.XPathSelectElements("one:Notebook", notebooks.Xnm))
                {
                    result.Add((string)notebook.Attribute("ID"), (string)notebook.Attribute("name"));
                }
            }

            return result;
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            List<string> selectedNotebooks = GetSelectedNotebooks();
            if (selectedNotebooks.Count == 0)
            {
                lblError.Text = BibleCommon.Resources.Constants.NoteLinkerNoElementSelected;
                lblError.Visible = true;                
            }
            else
            {
                Helper.SaveSelectedNotebooksIds(selectedNotebooks);
                this.DialogResult = System.Windows.Forms.DialogResult.OK;
                this.Close();
            }
        }

        private List<string> GetSelectedNotebooks()
        {
            List<string> result = new List<string>();

            foreach (CheckBox chk in pnMain.Controls)
            {
                if (chk.Checked)
                    result.Add(chk.Name);
            }

            return result;
        }

        private void SelectNoteBooksForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            _oneNoteApp = null;
        }
    }
}
