using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using BibleCommon.Helpers;
using System.Xml;
using System.Xml.Linq;
using Microsoft.Office.Interop.OneNote;
using System.Xml.XPath;
using BibleCommon;

namespace BibleConfigurator
{
    public partial class NotebookParametersForm : Form
    {
        private const string BibleSectionGroupDefaultName = "Библия";
        private const string BibleCommentsSectionGroupDefaultName = "Комментарии к Библии";
        private const string BibleStudySectionGroupDefaultName = "Изучение Библии";

        private Microsoft.Office.Interop.OneNote.Application _oneNoteApp;
        private string _notebookName;

        public Dictionary<string, string> RenamedSections = new Dictionary<string, string>();

        public NotebookParametersForm(Microsoft.Office.Interop.OneNote.Application oneNoteApp, string notebookName)
        {
            _oneNoteApp = oneNoteApp;
            _notebookName = notebookName;
            InitializeComponent();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {

        }

        private void NotebookParametersForm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                Close();
        }

        private void NotebookParametersForm_Load(object sender, EventArgs e)
        {
            Dictionary<string, string> allSectionGroups = GetAllSectionGroups();

            cbBibleSection.DataSource = allSectionGroups.Keys.ToList();
            cbBibleCommentsSection.DataSource = allSectionGroups.Keys.ToList();
            cbBibleStudySection.DataSource = allSectionGroups.Keys.ToList();

            cbBibleSection.SelectedItem = !string.IsNullOrEmpty(Settings.Default.SectionGroupName_Bible) ?
                Settings.Default.SectionGroupName_Bible : BibleSectionGroupDefaultName;
            cbBibleCommentsSection.SelectedItem = !string.IsNullOrEmpty(Settings.Default.SectionGroupName_BibleComments) ?
                Settings.Default.SectionGroupName_BibleComments : BibleCommentsSectionGroupDefaultName;
            cbBibleStudySection.SelectedItem = !string.IsNullOrEmpty(Settings.Default.SectionGroupName_BibleStudy) ?
                Settings.Default.SectionGroupName_BibleStudy : BibleStudySectionGroupDefaultName;
        }

        private Dictionary<string, string> GetAllSectionGroups()
        {
            Dictionary<string, string> result = new Dictionary<string, string>();

            string notebookId = OneNoteUtils.GetNotebookId(_oneNoteApp, _notebookName);

            if (!string.IsNullOrEmpty(notebookId))
            {
                string xml;
                XmlNamespaceManager xnm;
                _oneNoteApp.GetHierarchy(notebookId, HierarchyScope.hsSections, out xml);
                XDocument notebookDoc = OneNoteUtils.GetXDocument(xml, out xnm);

                foreach (XElement sectionGroup in notebookDoc.Root.XPathSelectElements("one:SectionGroup", xnm))
                {
                    string name = (string)sectionGroup.Attribute("name");
                    string id = (string)sectionGroup.Attribute("ID");

                    result.Add(name, id);
                }
            }
            else
            {
                Logger.LogError(string.Format("Не найдена записная книжка '{0}'.", _notebookName));
                this.DialogResult = System.Windows.Forms.DialogResult.None;
                Close();
            }

            return result;
        }
    }
}
