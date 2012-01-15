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

namespace BibleConfigurator
{
    public partial class MainForm : Form
    {
        private const string SingleNotebookDefaultName = "Holy Bible";
        private const string SingleNotebookTemplateFileName = "Holy Bible.onepkg";
        private const string BibleNotebookTemplateFileName = "Библия.onepkg";
        private const string BibleCommentsNotebookTemplateFileName = "Комментарии к Библии.onepkg";
        private const string BibleStudyNotebookTemplateFileName = "Изучение Библии.onepkg";
        private const string TemplatesDirectory = "OneNoteTemplates";

        private Microsoft.Office.Interop.OneNote.Application OneNoteApp = new Microsoft.Office.Interop.OneNote.Application();

        private string SingleNotebookFromTemplatePath { get; set; }
        private string BibleNotebookFromTemplatePath { get; set; }
        private string BibleCommentsNotebookFromTemplatePath { get; set; }
        private string BibleStudyNotebookFromTemplatePath { get; set; }

        public MainForm()
        {
            InitializeComponent();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if (rbSingleNotebook.Checked)
            {
                Settings.Default.NotebookName_Bible = string.Empty;
                Settings.Default.NotebookName_BibleComments = string.Empty;
                Settings.Default.NotebookName_BibleStudy = string.Empty;

                if (chkCreateSingleNotebookFromTemplate.Checked)
                {
                    CreateNotebookFromTemplate(SingleNotebookTemplateFileName, SingleNotebookFromTemplatePath);
                    Settings.Default.NotebookName_Single = Path.GetFileNameWithoutExtension(SingleNotebookTemplateFileName);
                }
                else
                    Settings.Default.NotebookName_Single = cbSingleNotebook.SelectedText;
            }
            else
            {
                Settings.Default.NotebookName_Single = string.Empty;
                Settings.Default.SectionGroupName_Bible = string.Empty;
                Settings.Default.SectionGroupName_BibleComments = string.Empty;
                Settings.Default.SectionGroupName_BibleStudy = string.Empty;

                if (chkCreateBibleNotebookFromTemplate.Checked)
                {
                    CreateNotebookFromTemplate(BibleNotebookTemplateFileName, BibleNotebookFromTemplatePath);
                    Settings.Default.NotebookName_Bible = Path.GetFileNameWithoutExtension(BibleNotebookTemplateFileName);
                }
                else
                    Settings.Default.NotebookName_Bible = cbBibleNotebook.SelectedText;

                if (chkCreateBibleCommentsNotebookFromTemplate.Checked)
                {
                    CreateNotebookFromTemplate(BibleCommentsNotebookTemplateFileName, BibleCommentsNotebookFromTemplatePath);
                    Settings.Default.NotebookName_BibleComments = Path.GetFileNameWithoutExtension(BibleCommentsNotebookTemplateFileName);
                }
                else
                    Settings.Default.NotebookName_BibleComments = cbBibleCommentsNotebook.SelectedText;

                if (chkCreateBibleStudyNotebookFromTemplate.Checked)
                {
                    CreateNotebookFromTemplate(BibleStudyNotebookTemplateFileName, BibleStudyNotebookFromTemplatePath);
                    Settings.Default.NotebookName_BibleStudy = Path.GetFileNameWithoutExtension(BibleStudyNotebookTemplateFileName);
                }
                else
                    Settings.Default.NotebookName_BibleStudy = cbBibleStudyNotebook.SelectedText;
            }
        }

        private void CreateNotebookFromTemplate(string notebookTemplateFileName, string notebookFromTemplatePath)
        {
            string s;
            OneNoteApp.OpenPackage(Path.Combine(Path.Combine(Path.GetPathRoot(Utils.GetCurrentDirectory()), TemplatesDirectory), notebookTemplateFileName),
                Path.Combine(notebookFromTemplatePath, Path.GetFileNameWithoutExtension(notebookTemplateFileName)), out s);

            string[] files = Directory.GetFiles(s, "*.onetoc2", SearchOption.TopDirectoryOnly);
            if (files.Length > 0)
                Process.Start(files[0]);
            else
                throw new Exception(string.Format("Ошибка при открытии записной книжки '{0}'", notebookTemplateFileName));
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            rbSingleNotebook.Checked = (Settings.Default.NotebookName_Bible == Settings.Default.NotebookName_BibleComments)
                                    && (Settings.Default.NotebookName_Bible == Settings.Default.NotebookName_BibleStudy);

            rbMultiNotebook.Checked = !rbSingleNotebook.Checked;
            rbMultiNotebook_CheckedChanged(this, null);


            Dictionary<string, string> notebooks = GetNotebooks();

            cbSingleNotebook.DataSource = notebooks.Values.ToList();
            cbSingleNotebook.SelectedItem = SingleNotebookDefaultName;
            cbBibleNotebook.DataSource = notebooks.Values.ToList();
            cbBibleNotebook.SelectedItem = Settings.Default.NotebookName_Bible;
            cbBibleCommentsNotebook.DataSource = notebooks.Values.ToList();
            cbBibleCommentsNotebook.SelectedItem = Settings.Default.NotebookName_BibleComments;
            cbBibleStudyNotebook.DataSource = notebooks.Values.ToList();
            cbBibleStudyNotebook.SelectedItem = Settings.Default.NotebookName_BibleStudy;

            string[] directories = Directory.GetDirectories(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), 
                                        "*OneNote*", SearchOption.TopDirectoryOnly);
            if (directories.Length > 0)
                folderBrowserDialog.SelectedPath = directories[0];

            folderBrowserDialog.RootFolder = Environment.SpecialFolder.MyDocuments;
        }

        public Dictionary<string, string> GetNotebooks()
        {
            Dictionary<string, string> result = new Dictionary<string, string>();

            string xml;
            XmlNamespaceManager xnm;
            OneNoteApp.GetHierarchy(null, HierarchyScope.hsNotebooks, out xml);
            XDocument doc = OneNoteUtils.GetXDocument(xml, out xnm);
            foreach (XElement notebook in doc.Root.XPathSelectElements("one:Notebook", xnm))
            {
                result.Add((string)notebook.Attribute("ID"), (string)notebook.Attribute("name"));
            }

            return result;
        }

        private void rbMultiNotebook_CheckedChanged(object sender, EventArgs e)
        {
            cbSingleNotebook.Enabled = rbSingleNotebook.Checked;
            btnSingleNotebookParameters.Enabled = rbSingleNotebook.Checked;
            chkCreateSingleNotebookFromTemplate.Enabled = rbSingleNotebook.Checked;

            cbBibleNotebook.Enabled = rbMultiNotebook.Checked;
            cbBibleCommentsNotebook.Enabled = rbMultiNotebook.Checked;
            cbBibleStudyNotebook.Enabled = rbMultiNotebook.Checked;
            chkCreateBibleNotebookFromTemplate.Enabled = rbMultiNotebook.Checked;
            chkCreateBibleCommentsNotebookFromTemplate.Enabled = rbMultiNotebook.Checked;
            chkCreateBibleStudyNotebookFromTemplate.Enabled = rbMultiNotebook.Checked;
        }

        private void chkCreateSingleNotebookFromTemplate_CheckedChanged(object sender, EventArgs e)
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


            cbSingleNotebook.Enabled = !chkCreateSingleNotebookFromTemplate.Checked;
        }

        private void chkCreateBibleNotebookFromTemplate_CheckedChanged(object sender, EventArgs e)
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

            cbBibleNotebook.Enabled = !chkCreateBibleNotebookFromTemplate.Checked;
        }

        private void chkCreateBibleCommentsNotebookFromTemplate_CheckedChanged(object sender, EventArgs e)
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

            cbBibleCommentsNotebook.Enabled = !chkCreateBibleCommentsNotebookFromTemplate.Checked;
        }

        private void chkCreateBibleStudyNotebookFromTemplate_CheckedChanged(object sender, EventArgs e)
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

            cbBibleStudyNotebook.Enabled = !chkCreateBibleStudyNotebookFromTemplate.Checked;
        }

        private void btnSingleNotebookParameters_Click(object sender, EventArgs e)
        {
            NotebookParametersForm notebookParametersForm = new NotebookParametersForm();
            notebookParametersForm.ShowDialog();
        }
    }
}
