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
        private Microsoft.Office.Interop.OneNote.Application _oneNoteApp;
        private string _notebookName;

        public Dictionary<SectionGroupType, SectionGroupInfo> OriginalSectionGroups { get; set; }
        public Dictionary<string, string> RenamedSectionGroups { get; set; }
        public Dictionary<SectionGroupType, string> GroupedSectionGroups { get; set; }
        

        public NotebookParametersForm(Microsoft.Office.Interop.OneNote.Application oneNoteApp, string notebookName)
        {
            _oneNoteApp = oneNoteApp;
            _notebookName = notebookName;
            InitializeComponent();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            GroupedSectionGroups = new Dictionary<SectionGroupType, string>();
            GroupedSectionGroups.Add(SectionGroupType.Bible, (string)cbBibleSection.SelectedItem);
            GroupedSectionGroups.Add(SectionGroupType.BibleComments, (string)cbBibleCommentsSection.SelectedItem);
            GroupedSectionGroups.Add(SectionGroupType.BibleStudy, (string)cbBibleStudySection.SelectedItem);
        }

        private void NotebookParametersForm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                Close();
        }

        private void NotebookParametersForm_Load(object sender, EventArgs e)
        {
            RenamedSectionGroups = new Dictionary<string, string>();

            try
            {
                OriginalSectionGroups = GetAllSectionGroups();               

                BindComboBox(cbBibleSection, OriginalSectionGroups[SectionGroupType.Bible]);

                BindComboBox(cbBibleCommentsSection, OriginalSectionGroups[SectionGroupType.BibleComments]);

                BindComboBox(cbBibleStudySection, OriginalSectionGroups[SectionGroupType.BibleStudy]);
            }
            catch (InvalidNotebookException)
            {
                Logger.LogError("Указана неподходящая записная книжка.");
                this.Close();
            }
        }

        private void BindComboBox(ComboBox cb, SectionGroupInfo sectionGroupInfo)
        {
            cb.Items.Add(sectionGroupInfo.OriginalName);
            cb.SelectedIndex = 0;
        }  

        private Dictionary<SectionGroupType, SectionGroupInfo> GetAllSectionGroups()
        {
            Dictionary<SectionGroupType, SectionGroupInfo> result = new Dictionary<SectionGroupType, SectionGroupInfo>();

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

                    if (NotebookChecker.ElementIsBible(sectionGroup, xnm) && !result.ContainsKey(SectionGroupType.Bible))
                        result.Add(SectionGroupType.Bible, new SectionGroupInfo() { Id = id, OriginalName = name, Type = SectionGroupType.Bible });
                    else if (NotebookChecker.ElementIsBibleComments(sectionGroup, xnm) && !result.ContainsKey(SectionGroupType.BibleComments))
                        result.Add(SectionGroupType.BibleComments, new SectionGroupInfo() { Id = id, OriginalName = name, Type = SectionGroupType.BibleComments });
                    else if (!result.ContainsKey(SectionGroupType.BibleStudy))
                        result.Add(SectionGroupType.BibleStudy, new SectionGroupInfo() { Id = id, OriginalName = name, Type = SectionGroupType.BibleStudy });
                    else
                        throw new InvalidNotebookException();
                }

                if (result.Count < 3)
                    throw new InvalidNotebookException();
            }
            else
            {
                Logger.LogError(string.Format("Не найдена записная книжка '{0}'.", _notebookName));
                this.DialogResult = System.Windows.Forms.DialogResult.None;
                Close();
            }

            return result;
        }

        private void btnBibleSectionRename_Click(object sender, EventArgs e)
        {
            TryToRenameSectionGroup(SectionGroupType.Bible, (string)cbBibleSection.SelectedItem);        
        }      

        private void btnBibleCommentsSectionRename_Click(object sender, EventArgs e)
        {
            TryToRenameSectionGroup(SectionGroupType.BibleComments, (string)cbBibleCommentsSection.SelectedItem);
        }

        private void btnBibleStudySectionRename_Click(object sender, EventArgs e)
        {
            TryToRenameSectionGroup(SectionGroupType.BibleStudy, (string)cbBibleStudySection.SelectedItem);
        }

        private void TryToRenameSectionGroup(SectionGroupType sectionGroupType, string originalSectionGroupName)
        {
            string result = CallRenameSectionGroupForm(originalSectionGroupName);
            if (!string.IsNullOrEmpty(result))
            {
                SectionGroupInfo sectionGroupInfo = OriginalSectionGroups[sectionGroupType];                
                sectionGroupInfo.NewName = result;
                RenamedSectionGroups.Add(sectionGroupInfo.Id, result);
                ChangeSectionGroupNameInComboBox(sectionGroupType, result);                
            }
        }

        private void ChangeSectionGroupNameInComboBox(SectionGroupType sectionGroupType, string result)
        {
            switch (sectionGroupType)
            {
                case SectionGroupType.Bible:
                    cbBibleSection.Items[0] = result;
                    break;
                case SectionGroupType.BibleComments:
                    cbBibleCommentsSection.Items[0] = result;
                    break;
                case SectionGroupType.BibleStudy:
                    cbBibleStudySection.Items[0] = result;
                    break;
            }
        }      

        private string CallRenameSectionGroupForm(string sectionGroupName)
        {
            string result = null;

            RenameSectionGroupsForm form = new RenameSectionGroupsForm(sectionGroupName);
            if (form.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                result = form.SectionGroupName;
            }

            return result;
        }
    }
}
