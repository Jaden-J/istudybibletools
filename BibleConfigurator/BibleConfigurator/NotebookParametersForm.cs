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
using BibleCommon.Common;

namespace BibleConfigurator
{
    public partial class NotebookParametersForm : Form
    {
        private Microsoft.Office.Interop.OneNote.Application _oneNoteApp;
        private string _notebookId;

        public Dictionary<ContainerType, SectionGroupDTO> OriginalSectionGroups { get; set; }
        public Dictionary<string, string> RenamedSectionGroups { get; set; }
        public Dictionary<ContainerType, string> GroupedSectionGroups { get; set; }

        private bool _firstLoad = true;

        public NotebookParametersForm(string notebookId)
        {
            _oneNoteApp = OneNoteUtils.CreateOneNoteAppSafe();
            _notebookId = notebookId;
            InitializeComponent();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            GroupedSectionGroups = new Dictionary<ContainerType, string>();
            GroupedSectionGroups.Add(ContainerType.Bible, OriginalSectionGroups[ContainerType.Bible].Id);
            GroupedSectionGroups.Add(ContainerType.BibleComments, OriginalSectionGroups[ContainerType.BibleComments].Id);
            GroupedSectionGroups.Add(ContainerType.BibleStudy, OriginalSectionGroups[ContainerType.BibleStudy].Id);            
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
                OneNoteProxy.HierarchyElement notebook = null;

                try
                {
                    notebook = OneNoteProxy.Instance.GetHierarchy(ref _oneNoteApp, _notebookId, HierarchyScope.hsSections, true);

                    OriginalSectionGroups = GetAllSectionGroups(notebook);

                    BindComboBox(cbBibleSection, OriginalSectionGroups[ContainerType.Bible]);                    

                    BindComboBox(cbBibleStudySection, OriginalSectionGroups[ContainerType.BibleStudy]);

                    BindComboBox(cbBibleCommentsSection, OriginalSectionGroups[ContainerType.BibleComments]);                    

                    _firstLoad = false;
                }
                catch (InvalidNotebookException)
                {
                    FormLogger.LogError(BibleCommon.Resources.Constants.ConfiguratorWrongNotebookSelected, 
                            notebook != null ? (string)notebook.Content.Root.Attribute("name") : null, 
                            ContainerType.Single);
                    this.Close();
                }
            }
        }

        private void BindComboBox(ComboBox cb, SectionGroupDTO sectionGroupInfo)
        {
            cb.Items.Add(sectionGroupInfo.OriginalName);
            cb.SelectedIndex = 0;
        }

        private Dictionary<ContainerType, SectionGroupDTO> GetAllSectionGroups(OneNoteProxy.HierarchyElement notebook)
        {
            Dictionary<ContainerType, SectionGroupDTO> result = new Dictionary<ContainerType, SectionGroupDTO>();
            
            var module = ModulesManager.GetCurrentModuleInfo();

            foreach (XElement sectionGroup in notebook.Content.Root.XPathSelectElements("one:SectionGroup", notebook.Xnm).Where(sg => !OneNoteUtils.IsRecycleBin(sg)))
            {
                string name = (string)sectionGroup.Attribute("name");
                string id = (string)sectionGroup.Attribute("ID");

                if (NotebookChecker.ElementIsBible(module, sectionGroup, notebook.Xnm) && !result.ContainsKey(ContainerType.Bible))
                    result.Add(ContainerType.Bible, new SectionGroupDTO() { Id = id, OriginalName = name, Type = ContainerType.Bible });
                else if (NotebookChecker.ElementIsBibleComments(module, sectionGroup, notebook.Xnm) && !result.ContainsKey(ContainerType.BibleComments))
                    result.Add(ContainerType.BibleComments, new SectionGroupDTO() { Id = id, OriginalName = name, Type = ContainerType.BibleComments });                
                else if (!result.ContainsKey(ContainerType.BibleStudy))
                    result.Add(ContainerType.BibleStudy, new SectionGroupDTO() { Id = id, OriginalName = name, Type = ContainerType.BibleStudy });
                else
                    throw new InvalidNotebookException();
            }

            if (result.Count < 3)
                throw new InvalidNotebookException();

            return result;
        }

        private void btnBibleSectionRename_Click(object sender, EventArgs e)
        {
            TryToRenameSectionGroup(ContainerType.Bible, (string)cbBibleSection.SelectedItem);        
        }      

        private void btnBibleCommentsSectionRename_Click(object sender, EventArgs e)
        {
            TryToRenameSectionGroup(ContainerType.BibleComments, (string)cbBibleCommentsSection.SelectedItem);
        }

        private void btnBibleStudySectionRename_Click(object sender, EventArgs e)
        {
            TryToRenameSectionGroup(ContainerType.BibleStudy, (string)cbBibleStudySection.SelectedItem);
        }

        private void TryToRenameSectionGroup(ContainerType sectionGroupType, string originalSectionGroupName)
        {
            string result = CallRenameSectionGroupForm(originalSectionGroupName);
            if (!string.IsNullOrEmpty(result))
            {
                SectionGroupDTO sectionGroupInfo = OriginalSectionGroups[sectionGroupType];                
                sectionGroupInfo.NewName = result;

                if (!RenamedSectionGroups.ContainsKey(sectionGroupInfo.Id))
                    RenamedSectionGroups.Add(sectionGroupInfo.Id, result);
                else
                    RenamedSectionGroups[sectionGroupInfo.Id] = result;

                ChangeSectionGroupNameInComboBox(sectionGroupType, result);                
            }
        }

        private void ChangeSectionGroupNameInComboBox(ContainerType sectionGroupType, string result)
        {
            switch (sectionGroupType)
            {
                case ContainerType.Bible:
                    cbBibleSection.Items[0] = result;
                    break;
                case ContainerType.BibleComments:
                    cbBibleCommentsSection.Items[0] = result;
                    break;
                case ContainerType.BibleStudy:
                    cbBibleStudySection.Items[0] = result;
                    break;
            }
        }      

        private string CallRenameSectionGroupForm(string sectionGroupName)
        {
            string result = null;

            using (RenameSectionGroupsForm form = new RenameSectionGroupsForm(sectionGroupName))
            {
                if (form.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    result = form.SectionGroupName;
                }
            }

            return result;
        }

        private void NotebookParametersForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            OneNoteUtils.ReleaseOneNoteApp(ref _oneNoteApp);
        }       
    }
}
