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
using BibleCommon.Services;

namespace BibleConfigurator
{
    public partial class NotebookParametersForm : Form
    {
        private Microsoft.Office.Interop.OneNote.Application _oneNoteApp;
        private string _notebookId;

        public Dictionary<SectionGroupType, SectionGroupInfo> OriginalSectionGroups { get; set; }
        public Dictionary<string, string> RenamedSectionGroups { get; set; }
        public Dictionary<SectionGroupType, string> GroupedSectionGroups { get; set; }

        private bool _firstLoad = true;

        public NotebookParametersForm(Microsoft.Office.Interop.OneNote.Application oneNoteApp, string notebookId)
        {
            _oneNoteApp = oneNoteApp;
            _notebookId = notebookId;
            InitializeComponent();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            GroupedSectionGroups = new Dictionary<SectionGroupType, string>();
            GroupedSectionGroups.Add(SectionGroupType.Bible, OriginalSectionGroups[SectionGroupType.Bible].Id);
            GroupedSectionGroups.Add(SectionGroupType.BibleComments, OriginalSectionGroups[SectionGroupType.BibleComments].Id);
            GroupedSectionGroups.Add(SectionGroupType.BibleStudy, OriginalSectionGroups[SectionGroupType.BibleStudy].Id);
            GroupedSectionGroups.Add(SectionGroupType.BibleNotesPages, OriginalSectionGroups[SectionGroupType.BibleNotesPages].Id);
        }

        private void NotebookParametersForm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                Close();
        }

        private void NotebookParametersForm_Load(object sender, EventArgs e)
        {
            if (_firstLoad)
            {
                RenamedSectionGroups = new Dictionary<string, string>();

                try
                {
                    OriginalSectionGroups = GetAllSectionGroups();

                    BindComboBox(cbBibleSection, OriginalSectionGroups[SectionGroupType.Bible]);                    

                    BindComboBox(cbBibleStudySection, OriginalSectionGroups[SectionGroupType.BibleStudy]);

                    BindComboBox(cbBibleCommentsSection, OriginalSectionGroups[SectionGroupType.BibleComments]);

                    BindComboBox(cbBibleNotesPagesSection, OriginalSectionGroups[SectionGroupType.BibleNotesPages]);

                    _firstLoad = false;
                }
                catch (InvalidNotebookException)
                {
                    Logger.LogError("Указана неподходящая записная книжка.");
                    this.Close();
                }
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

            OneNoteProxy.HierarchyElement notebook = OneNoteProxy.Instance.GetHierarchy(_oneNoteApp, _notebookId, HierarchyScope.hsSections, true);                            

            foreach (XElement sectionGroup in notebook.Content.Root.XPathSelectElements("one:SectionGroup", notebook.Xnm).Where(sg => !OneNoteUtils.IsRecycleBin(sg)))
            {
                string name = (string)sectionGroup.Attribute("name");
                string id = (string)sectionGroup.Attribute("ID");

                if (NotebookChecker.ElementIsBible(sectionGroup, notebook.Xnm) && !result.ContainsKey(SectionGroupType.Bible))
                    result.Add(SectionGroupType.Bible, new SectionGroupInfo() { Id = id, OriginalName = name, Type = SectionGroupType.Bible });
                else if (NotebookChecker.ElementIsBibleComments(sectionGroup, notebook.Xnm) && !result.ContainsKey(SectionGroupType.BibleComments))
                    result.Add(SectionGroupType.BibleComments, new SectionGroupInfo() { Id = id, OriginalName = name, Type = SectionGroupType.BibleComments });
                else if (NotebookChecker.ElementIsBibleComments(sectionGroup, notebook.Xnm) && !result.ContainsKey(SectionGroupType.BibleNotesPages))
                    result.Add(SectionGroupType.BibleNotesPages, new SectionGroupInfo() { Id = id, OriginalName = name, Type = SectionGroupType.BibleNotesPages});
                else if (!result.ContainsKey(SectionGroupType.BibleStudy))
                    result.Add(SectionGroupType.BibleStudy, new SectionGroupInfo() { Id = id, OriginalName = name, Type = SectionGroupType.BibleStudy });
                else
                    throw new InvalidNotebookException();
            }

            if (result.Count < 3)
                throw new InvalidNotebookException();

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

        private void btnNotesPagesSectionRename_Click(object sender, EventArgs e)
        {
            TryToRenameSectionGroup(SectionGroupType.BibleNotesPages, (string)cbBibleNotesPagesSection.SelectedItem);
        }

        private void TryToRenameSectionGroup(SectionGroupType sectionGroupType, string originalSectionGroupName)
        {
            string result = CallRenameSectionGroupForm(originalSectionGroupName);
            if (!string.IsNullOrEmpty(result))
            {
                SectionGroupInfo sectionGroupInfo = OriginalSectionGroups[sectionGroupType];                
                sectionGroupInfo.NewName = result;

                if (!RenamedSectionGroups.ContainsKey(sectionGroupInfo.Id))
                    RenamedSectionGroups.Add(sectionGroupInfo.Id, result);
                else
                    RenamedSectionGroups[sectionGroupInfo.Id] = result;

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
                case SectionGroupType.BibleNotesPages:
                    cbBibleNotesPagesSection.Items[0] = result;
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
