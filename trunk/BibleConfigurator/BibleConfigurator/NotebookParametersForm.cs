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

        public Dictionary<string, string> OriginalSectionGroups { get; set; }
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
            OriginalSectionGroups = GetAllSectionGroups();
            RenamedSectionGroups = new Dictionary<string, string>();

            BindComboBox(cbBibleSection, OriginalSectionGroups.Values.ToList(), !string.IsNullOrEmpty(Settings.Default.SectionGroupName_Bible) 
                ? Settings.Default.SectionGroupName_Bible : Consts.BibleSectionGroupDefaultName);                

            BindComboBox(cbBibleCommentsSection, OriginalSectionGroups.Values.ToList(), !string.IsNullOrEmpty(Settings.Default.SectionGroupName_BibleComments)
                ? Settings.Default.SectionGroupName_BibleComments : Consts.BibleCommentsSectionGroupDefaultName);

            BindComboBox(cbBibleStudySection, OriginalSectionGroups.Values.ToList(), !string.IsNullOrEmpty(Settings.Default.SectionGroupName_BibleStudy)
                ? Settings.Default.SectionGroupName_BibleStudy : Consts.BibleStudySectionGroupDefaultName);
        }

        private void BindComboBox(ComboBox cb, List<string> dataSource, string selectedItem)
        {   
            cb.DataSource = dataSource;
            if (!string.IsNullOrEmpty(selectedItem))
                cb.SelectedItem = selectedItem;
        }

        private void BindComboBoxByOriginalSectionGroup(ComboBox cb, string selectedSectionGroupId, bool selectItem)
        {            
            object oldSelectedItem = cb.SelectedItem;
            cb.DataSource = OriginalSectionGroups.Values.ToList();
            if (selectItem)
                cb.SelectedItem = OriginalSectionGroups[selectedSectionGroupId];
            else
                cb.SelectedItem = oldSelectedItem;
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

                    result.Add(id, name);                  
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

        private void btnBibleSectionRename_Click(object sender, EventArgs e)
        {
            TryToRenameSectionGroup((string)cbBibleSection.SelectedItem);        
        }      

        private void btnBibleCommentsSectionRename_Click(object sender, EventArgs e)
        {
            TryToRenameSectionGroup((string)cbBibleCommentsSection.SelectedItem);
        }

        private void btnBibleStudySectionRename_Click(object sender, EventArgs e)
        {
            TryToRenameSectionGroup((string)cbBibleStudySection.SelectedItem);
        }

        private void ChangeSectionGroupNameInAllComboBoxes(string sectionGroupId, string originalSectionGroupName)
        {
            ChangeSectionGroupName(cbBibleSection, sectionGroupId, originalSectionGroupName);
            ChangeSectionGroupName(cbBibleCommentsSection, sectionGroupId, originalSectionGroupName);
            ChangeSectionGroupName(cbBibleStudySection, sectionGroupId, originalSectionGroupName);
        }

        private void ChangeSectionGroupName(ComboBox cb, string sectionGroupId, string originalSectionGroupName)
        {            
            BindComboBoxByOriginalSectionGroup(cb, sectionGroupId, (string)cb.SelectedItem == originalSectionGroupName);            
        }

        private void TryToRenameSectionGroup(string originalSectionGroupName)
        {
            string result = CallRenameSectionGroupForm(originalSectionGroupName);
            if (!string.IsNullOrEmpty(result))
            {                
                string sectionGroupId = GetDictionaryKeyBaValue(OriginalSectionGroups, originalSectionGroupName);
                OriginalSectionGroups[sectionGroupId] = result;
                RenamedSectionGroups.Add(sectionGroupId, result);
                ChangeSectionGroupNameInAllComboBoxes(sectionGroupId, originalSectionGroupName);                
            }
        }

        private string GetDictionaryKeyBaValue(Dictionary<string, string> dictionary, string value)
        {
            foreach (string key in dictionary.Keys)
            {
                if (dictionary[key] == value)
                    return key;
            }

            return null;
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
