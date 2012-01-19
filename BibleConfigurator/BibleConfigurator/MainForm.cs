using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using BibleCommon.Services;
using BibleCommon.Helpers;
using System.Xml;
using Microsoft.Office.Interop.OneNote;
using System.Xml.Linq;
using System.Xml.XPath;
using System.IO;
using System.Diagnostics;
using BibleCommon;
using System.Threading;

namespace BibleConfigurator
{
    public partial class MainForm : Form
    {
        private Microsoft.Office.Interop.OneNote.Application _oneNoteApp = new Microsoft.Office.Interop.OneNote.Application();

        private string SingleNotebookFromTemplatePath { get; set; }
        private string BibleNotebookFromTemplatePath { get; set; }
        private string BibleCommentsNotebookFromTemplatePath { get; set; }
        private string BibleStudyNotebookFromTemplatePath { get; set; }

        private bool _wasSearchedSectionGroupsInSingleNotebook = false;
        
        public Dictionary<string, string> RenamedSectionGroups { get; set; }
        public bool ToRenameSectionGroups { get; set; }


       

        public MainForm()
        {
            InitializeComponent();                     
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            btnOK.Enabled = false;

            Logger.Initialize();       

            if (rbSingleNotebook.Checked)
            {
                Settings.Default.NotebookName_Bible = string.Empty;
                Settings.Default.NotebookName_BibleComments = string.Empty;
                Settings.Default.NotebookName_BibleStudy = string.Empty;

                if (chkCreateSingleNotebookFromTemplate.Checked)
                {
                    string notebookName = CreateNotebookFromTemplate(Consts.SingleNotebookTemplateFileName, SingleNotebookFromTemplatePath);
                    if (!string.IsNullOrEmpty(notebookName))
                    {
                        Settings.Default.NotebookName_Single = notebookName;

                        Settings.Default.SectionGroupName_Bible = Consts.BibleSectionGroupDefaultName;
                        Settings.Default.SectionGroupName_BibleComments = Consts.BibleCommentsSectionGroupDefaultName;
                        Settings.Default.SectionGroupName_BibleStudy = Consts.BibleStudySectionGroupDefaultName;
                    }
                }
                else
                {
                    string notebookName = (string)cbSingleNotebook.SelectedItem;
                    if (NotebookChecker.CheckNotebook(_oneNoteApp, notebookName, NotebookType.Single))
                    {
                        Settings.Default.NotebookName_Single = notebookName;

                        if (ToRenameSectionGroups)
                            RenameSectionGroupsForm(notebookName, RenamedSectionGroups);

                        if (!_wasSearchedSectionGroupsInSingleNotebook)
                        {
                            try
                            {
                                SearchForCorrespondenceSectionGroups(notebookName);
                            }
                            catch (InvalidNotebookException)
                            {
                                Logger.LogError("Указана неподходящая записная книжка.");
                            }
                        }
                    }
                }
            }
            else
            {
                Settings.Default.NotebookName_Single = string.Empty;
                Settings.Default.SectionGroupName_Bible = string.Empty;
                Settings.Default.SectionGroupName_BibleComments = string.Empty;
                Settings.Default.SectionGroupName_BibleStudy = string.Empty;

                if (chkCreateBibleStudyNotebookFromTemplate.Checked)
                {
                    string notebookName = CreateNotebookFromTemplate(Consts.BibleStudyNotebookTemplateFileName, BibleStudyNotebookFromTemplatePath);
                    if (!string.IsNullOrEmpty(notebookName))
                    {
                        Settings.Default.NotebookName_BibleStudy = notebookName;
                        Thread.Sleep(3000);  // чтоб точно OneNote отработал
                    }
                }
                else
                {
                    if (NotebookChecker.CheckNotebook(_oneNoteApp, (string)cbBibleStudyNotebook.SelectedItem, NotebookType.BibleStudy))
                        Settings.Default.NotebookName_BibleStudy = (string)cbBibleStudyNotebook.SelectedItem;
                }

                if (chkCreateBibleCommentsNotebookFromTemplate.Checked)
                {
                    string notebookName = CreateNotebookFromTemplate(Consts.BibleCommentsNotebookTemplateFileName, BibleCommentsNotebookFromTemplatePath);
                    if (!string.IsNullOrEmpty(notebookName))
                    {
                        Settings.Default.NotebookName_BibleComments = notebookName;
                        Thread.Sleep(3000);  // чтоб точно OneNote отработал
                    }
                }
                else
                {
                    if (NotebookChecker.CheckNotebook(_oneNoteApp, (string)cbBibleCommentsNotebook.SelectedItem, NotebookType.BibleComments))
                        Settings.Default.NotebookName_BibleComments = (string)cbBibleCommentsNotebook.SelectedItem;
                }

                if (chkCreateBibleNotebookFromTemplate.Checked)  // записную книжку для Библии создаём в самом конце, так как она дольше всех создаётся
                {
                    string notebookName = CreateNotebookFromTemplate(Consts.BibleNotebookTemplateFileName, BibleNotebookFromTemplatePath);
                    if (!string.IsNullOrEmpty(notebookName))
                        Settings.Default.NotebookName_Bible = notebookName;
                }
                else
                {
                    if (NotebookChecker.CheckNotebook(_oneNoteApp, (string)cbBibleNotebook.SelectedItem, NotebookType.Bible))
                        Settings.Default.NotebookName_Bible = (string)cbBibleNotebook.SelectedItem;
                }
            }

            if (!Logger.WasErrorLogged)
            {
                if (chkDefaultPageNameParameters.Checked)
                {
                    Settings.Default.PageName_DefaultBookOverview = Consts.PageNameDefaultBookOverview;
                    Settings.Default.PageName_Notes = Consts.PageNameNotes;
                    Settings.Default.PageName_DefaultDescription = Consts.PageNameDefaultDescription;
                }
                else
                {
                    if (!string.IsNullOrEmpty(tbBookOverviewName.Text))
                        Settings.Default.PageName_DefaultBookOverview = tbBookOverviewName.Text;

                    if (!string.IsNullOrEmpty(tbNotesPageName.Text))
                        Settings.Default.PageName_Notes = tbNotesPageName.Text;

                    if (!string.IsNullOrEmpty(tbPageDescriptionName.Text))
                        Settings.Default.PageName_DefaultDescription = tbPageDescriptionName.Text;
                }

                Settings.Default.Save();
                Close();
            }
            else
                btnOK.Enabled = true;
        }

        private void SearchForCorrespondenceSectionGroups(string notebookName)
        {
            string notebookId = OneNoteUtils.GetNotebookId(_oneNoteApp, notebookName);

            if (!string.IsNullOrEmpty(notebookId))
            {
                string xml;
                XmlNamespaceManager xnm;
                _oneNoteApp.GetHierarchy(notebookId, HierarchyScope.hsSections, out xml);
                XDocument notebookDoc = OneNoteUtils.GetXDocument(xml, out xnm);

                List<SectionGroupType> sectionGroups = new List<SectionGroupType>();

                foreach (XElement sectionGroup in notebookDoc.Root.XPathSelectElements("one:SectionGroup", xnm))
                {
                    string name = (string)sectionGroup.Attribute("name");

                    if (NotebookChecker.ElementIsBible(sectionGroup, xnm) && !sectionGroups.Contains(SectionGroupType.Bible))
                    {
                        Settings.Default.SectionGroupName_Bible = name;
                        sectionGroups.Add(SectionGroupType.Bible);
                    }
                    else if (NotebookChecker.ElementIsBibleComments(sectionGroup, xnm) && !sectionGroups.Contains(SectionGroupType.BibleComments))
                    {
                        Settings.Default.SectionGroupName_BibleComments = name;
                        sectionGroups.Add(SectionGroupType.BibleComments);
                    }
                    else if (!sectionGroups.Contains(SectionGroupType.BibleStudy))
                    {
                        Settings.Default.SectionGroupName_BibleStudy = name;
                        sectionGroups.Add(SectionGroupType.BibleStudy);
                    }
                    else
                        throw new InvalidNotebookException();
                }

                if (sectionGroups.Count < 3)
                    throw new InvalidNotebookException();
            }
            else
            {
                Logger.LogError(string.Format("Не найдена записная книжка '{0}'.", notebookName));
            }
        }

        private void RenameSectionGroupsForm(string notebookName, Dictionary<string, string> renamedSectionGroups)
        {
            string notebookId = OneNoteUtils.GetNotebookId(_oneNoteApp, notebookName);

            if (!string.IsNullOrEmpty(notebookId))
            {
                string xml;
                XmlNamespaceManager xnm;
                _oneNoteApp.GetHierarchy(notebookId, HierarchyScope.hsSections, out xml);
                XDocument notebookDoc = OneNoteUtils.GetXDocument(xml, out xnm);

                foreach (string sectionGroupId in renamedSectionGroups.Keys)
                {
                    XElement sectionGroup = notebookDoc.Root.XPathSelectElement(string.Format("one:SectionGroup[@ID='{0}']", sectionGroupId), xnm);

                    if (sectionGroup != null)
                    {
                        sectionGroup.SetAttributeValue("name", renamedSectionGroups[sectionGroupId]);
                    }
                    else
                        Logger.LogError(string.Format("Не найдена группа разделов '{0}'.", sectionGroupId));
                }

                _oneNoteApp.UpdateHierarchy(notebookDoc.ToString());
            }
            else
                Logger.LogError(string.Format("Не найдена записная книжка '{0}'.", notebookName));
        }

        private string CreateNotebookFromTemplate(string notebookTemplateFileName, string notebookFromTemplatePath)
        {
            string s;
            string packageFilePath = Path.Combine(Path.Combine(Path.GetDirectoryName(Utils.GetCurrentDirectory()), Consts.TemplatesDirectory), notebookTemplateFileName);

            if (File.Exists(packageFilePath))
            {
                string folderPath = Path.Combine(notebookFromTemplatePath, Path.GetFileNameWithoutExtension(notebookTemplateFileName));                

                folderPath = GetNewDirectoryPath(folderPath);

                if (!string.IsNullOrEmpty(folderPath))
                {
                    _oneNoteApp.OpenPackage(packageFilePath, folderPath, out s);

                    string[] files = Directory.GetFiles(s, "*.onetoc2", SearchOption.TopDirectoryOnly);
                    if (files.Length > 0)
                        Process.Start(files[0]);
                    else
                        Logger.LogError(string.Format("Ошибка при открытии записной книжки '{0}'.", notebookTemplateFileName));

                    return Path.GetFileNameWithoutExtension(folderPath);
                }
                else
                    Logger.LogError("Не удаётся создать записную книжку. Выберите другую папку.");
            }
            else
                Logger.LogError(string.Format("Не найден шаблон записной книжки по адресу '{0}'.", packageFilePath));

            return string.Empty;
        }

        private string GetNewDirectoryPath(string folderPath)
        {
            string result = folderPath;
            for (int i = 0; i < 100; i++)
            {
                result = folderPath + (i > 0 ? " (" + i.ToString() + ")" : string.Empty);

                if (!Directory.Exists(result))
                    return result;
            }

            return string.Empty;
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            PrepareFolderBrowser();
            SetNotebooksDefaultPaths();

            rbSingleNotebook.Checked = !string.IsNullOrEmpty(Settings.Default.NotebookName_Single);

            rbMultiNotebook.Checked = !rbSingleNotebook.Checked;
            rbMultiNotebook_CheckedChanged(this, null);


            Dictionary<string, string> notebooks = GetNotebooks();

            cbSingleNotebook.DataSource = notebooks.Keys.ToList();
            cbBibleNotebook.DataSource = notebooks.Keys.ToList();
            cbBibleCommentsNotebook.DataSource = notebooks.Keys.ToList();
            cbBibleStudyNotebook.DataSource = notebooks.Keys.ToList();

            string singleNotebookName = !string.IsNullOrEmpty(Settings.Default.NotebookName_Single) ?
                Settings.Default.NotebookName_Single : Consts.SingleNotebookDefaultName;
            if (cbSingleNotebook.Items.Contains(singleNotebookName))
                cbSingleNotebook.SelectedItem = singleNotebookName;
            else
                chkCreateSingleNotebookFromTemplate.Checked = true;

            string bibleNotebookName = !string.IsNullOrEmpty(Settings.Default.NotebookName_Bible) ?
                Settings.Default.NotebookName_Bible : Consts.BibleNotebookDefaultName;
            if (cbBibleNotebook.Items.Contains(bibleNotebookName))
                cbBibleNotebook.SelectedItem = bibleNotebookName;
            else
                chkCreateBibleNotebookFromTemplate.Checked = true;

            string bibleCommentsNotebookName = !string.IsNullOrEmpty(Settings.Default.NotebookName_BibleComments) ?
                Settings.Default.NotebookName_BibleComments : Consts.BibleCommentsNotebookDefaultName;
            if (cbBibleCommentsNotebook.Items.Contains(bibleCommentsNotebookName))
                cbBibleCommentsNotebook.SelectedItem = bibleCommentsNotebookName;
            else
                chkCreateBibleCommentsNotebookFromTemplate.Checked = true;

            string bibleStudyNotebookName = !string.IsNullOrEmpty(Settings.Default.NotebookName_BibleStudy) ?
                Settings.Default.NotebookName_BibleStudy : Consts.BibleStudyNotebookDefaultName;
            if (cbBibleStudyNotebook.Items.Contains(bibleStudyNotebookName))
                cbBibleStudyNotebook.SelectedItem = bibleStudyNotebookName;
            else
                chkCreateBibleStudyNotebookFromTemplate.Checked = true;

            tbBookOverviewName.Text = Settings.Default.PageName_DefaultBookOverview;
            tbNotesPageName.Text = Settings.Default.PageName_Notes;
            tbPageDescriptionName.Text = Settings.Default.PageName_DefaultDescription;
        }

        private void SetNotebooksDefaultPaths()
        {
            // по дефолту пути такие
            SingleNotebookFromTemplatePath = folderBrowserDialog.SelectedPath;
            BibleNotebookFromTemplatePath = folderBrowserDialog.SelectedPath;
            BibleCommentsNotebookFromTemplatePath = folderBrowserDialog.SelectedPath;
            BibleStudyNotebookFromTemplatePath = folderBrowserDialog.SelectedPath;
        }

        private void PrepareFolderBrowser()
        {
            string myDocumentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string[] directories = Directory.GetDirectories(myDocumentsPath, "*OneNote*", SearchOption.TopDirectoryOnly);
            if (directories.Length > 0)
                folderBrowserDialog.SelectedPath = directories[0];
            else
                folderBrowserDialog.SelectedPath = myDocumentsPath;            

            folderBrowserDialog.Description = "Укажите расположение записной книжки";
            folderBrowserDialog.ShowNewFolderButton = true;            
        }

        public Dictionary<string, string> GetNotebooks()
        {
            Dictionary<string, string> result = new Dictionary<string, string>();

            string xml;
            XmlNamespaceManager xnm;
            _oneNoteApp.GetHierarchy(null, HierarchyScope.hsNotebooks, out xml);
            XDocument doc = OneNoteUtils.GetXDocument(xml, out xnm);
            foreach (XElement notebook in doc.Root.XPathSelectElements("one:Notebook", xnm))
            {
                string name = (string)notebook.Attribute("name");
                string id = (string)notebook.Attribute("ID");
                result.Add(name, id);
            }

            return result;
        }

        private void rbMultiNotebook_CheckedChanged(object sender, EventArgs e)
        {
            cbSingleNotebook.Enabled = rbSingleNotebook.Checked;
            btnSingleNotebookParameters.Enabled = rbSingleNotebook.Checked;
            chkCreateSingleNotebookFromTemplate.Enabled = rbSingleNotebook.Checked;
            btnSingleNotebookParameters.Enabled = rbSingleNotebook.Checked;
            btnSingleNotebookSetPath.Enabled = rbSingleNotebook.Checked;

            cbBibleNotebook.Enabled = rbMultiNotebook.Checked;
            cbBibleCommentsNotebook.Enabled = rbMultiNotebook.Checked;
            cbBibleStudyNotebook.Enabled = rbMultiNotebook.Checked;
            chkCreateBibleNotebookFromTemplate.Enabled = rbMultiNotebook.Checked;
            chkCreateBibleCommentsNotebookFromTemplate.Enabled = rbMultiNotebook.Checked;
            chkCreateBibleStudyNotebookFromTemplate.Enabled = rbMultiNotebook.Checked;
            btnBibleNotebookSetPath.Enabled = rbMultiNotebook.Checked;
            btnBibleCommentsNotebookSetPath.Enabled = rbMultiNotebook.Checked;
            btnBibleStudyNotebookSetPath.Enabled = rbMultiNotebook.Checked;

            if (rbSingleNotebook.Checked)
            {
                chkCreateSingleNotebookFromTemplate_CheckedChanged(this, null);
            }
            else
            {
                chkCreateBibleNotebookFromTemplate_CheckedChanged(this, null);
                chkCreateBibleCommentsNotebookFromTemplate_CheckedChanged(this, null);
                chkCreateBibleStudyNotebookFromTemplate_CheckedChanged(this, null);
            }
        }

        private void chkCreateSingleNotebookFromTemplate_CheckedChanged(object sender, EventArgs e)
        {
            cbSingleNotebook.Enabled = !chkCreateSingleNotebookFromTemplate.Checked;
            btnSingleNotebookParameters.Enabled = !chkCreateSingleNotebookFromTemplate.Checked;
            btnSingleNotebookSetPath.Enabled = chkCreateSingleNotebookFromTemplate.Checked;
        }

        private void chkCreateBibleNotebookFromTemplate_CheckedChanged(object sender, EventArgs e)
        {
            cbBibleNotebook.Enabled = !chkCreateBibleNotebookFromTemplate.Checked;
            btnBibleNotebookSetPath.Enabled = chkCreateBibleNotebookFromTemplate.Checked;
        }

        private void chkCreateBibleCommentsNotebookFromTemplate_CheckedChanged(object sender, EventArgs e)
        {
            cbBibleCommentsNotebook.Enabled = !chkCreateBibleCommentsNotebookFromTemplate.Checked;
            btnBibleCommentsNotebookSetPath.Enabled = chkCreateBibleCommentsNotebookFromTemplate.Checked;
        }

        private void chkCreateBibleStudyNotebookFromTemplate_CheckedChanged(object sender, EventArgs e)
        {
            cbBibleStudyNotebook.Enabled = !chkCreateBibleStudyNotebookFromTemplate.Checked;
            btnBibleStudyNotebookSetPath.Enabled = chkCreateBibleStudyNotebookFromTemplate.Checked;
        }

        private void btnSingleNotebookParameters_Click(object sender, EventArgs e)
        {   
            if (!string.IsNullOrEmpty((string)cbSingleNotebook.SelectedItem))
            {
                string notebookName = (string)cbSingleNotebook.SelectedItem;
                if (NotebookChecker.CheckNotebook(_oneNoteApp, notebookName, NotebookType.Single))
                {
                    NotebookParametersForm notebookParametersForm = new NotebookParametersForm(_oneNoteApp, notebookName);
                    if (notebookParametersForm.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        if (notebookParametersForm.RenamedSectionGroups.Count > 0)
                        {
                            ToRenameSectionGroups = true;
                            RenamedSectionGroups = notebookParametersForm.RenamedSectionGroups;
                        }
                        
                        Settings.Default.SectionGroupName_Bible = notebookParametersForm.GroupedSectionGroups[SectionGroupType.Bible];                        
                        Settings.Default.SectionGroupName_BibleComments = notebookParametersForm.GroupedSectionGroups[SectionGroupType.BibleComments];                        
                        Settings.Default.SectionGroupName_BibleStudy = notebookParametersForm.GroupedSectionGroups[SectionGroupType.BibleStudy];

                        _wasSearchedSectionGroupsInSingleNotebook = true;  // нашли необходимые группы секций. 
                    }
                }
            }
            else
            {
                Logger.LogMessage("Не указана записная книжка.");
            }
        }

        private void btnSingleNotebookSetPath_Click(object sender, EventArgs e)
        {
            if (chkCreateSingleNotebookFromTemplate.Checked)
            {
                if (folderBrowserDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    SingleNotebookFromTemplatePath = folderBrowserDialog.SelectedPath;
                }
                else
                    chkCreateSingleNotebookFromTemplate.Checked = false;
            }
        }

        private void btnBibleNotebookSetPath_Click(object sender, EventArgs e)
        {
            if (chkCreateBibleNotebookFromTemplate.Checked)
            {
                if (folderBrowserDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    BibleNotebookFromTemplatePath = folderBrowserDialog.SelectedPath;
                }
                else
                    chkCreateBibleNotebookFromTemplate.Checked = false;
            }
        }

        private void btnBibleCommentsNotebookSetPath_Click(object sender, EventArgs e)
        {
            if (chkCreateBibleCommentsNotebookFromTemplate.Checked)
            {
                if (folderBrowserDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    BibleCommentsNotebookFromTemplatePath = folderBrowserDialog.SelectedPath;
                }
                else
                    chkCreateBibleCommentsNotebookFromTemplate.Checked = false;
            }
        }

        private void btnBibleStudyNotebookSetPath_Click(object sender, EventArgs e)
        {
            if (chkCreateBibleStudyNotebookFromTemplate.Checked)
            {
                if (folderBrowserDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    BibleStudyNotebookFromTemplatePath = folderBrowserDialog.SelectedPath;
                }
                else
                    chkCreateBibleStudyNotebookFromTemplate.Checked = false;
            }
        }

        private void chkDefaultPageNameParameters_CheckedChanged(object sender, EventArgs e)
        {
            tbPageDescriptionName.Enabled = !chkDefaultPageNameParameters.Checked;
            tbNotesPageName.Enabled = !chkDefaultPageNameParameters.Checked;
            tbBookOverviewName.Enabled = !chkDefaultPageNameParameters.Checked;
        }
    }
}
