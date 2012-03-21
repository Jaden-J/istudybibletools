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

namespace BibleNoteLinkerEx
{
    public partial class SelectNoteBooks : Form
    {
        private Microsoft.Office.Interop.OneNote.Application _oneNoteApp;

        public SelectNoteBooks(Microsoft.Office.Interop.OneNote.Application oneNoteApp)
        {
            InitializeComponent();
            _oneNoteApp = oneNoteApp;
        }

        private void SelectNoteBooks_Load(object sender, EventArgs e)
        {
            Dictionary<string, string> allNotebooks = GetAllNotebooks();

            int i = 0;
            foreach (string id in allNotebooks.Keys)
            {
                RenderSelectRow(id, allNotebooks[id], i++);
            }

            btnOk.Top = 25 * i + 20;
            this.Height = btnOk.Top + btnOk.Height + 35;
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
                //todo
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
    }
}
