using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using BibleConfigurator.ModuleConverter;
using BibleCommon.Helpers;
using BibleCommon.Common;
using System.IO;
using BibleCommon.Scheme;

namespace BibleConfigurator
{
    public partial class ZefaniaXmlConverterForm : Form
    {
        public const string LocaleStructureFilePath = "structure.xml";
        public const string LocaleBooksInfoFilePath = "books.xml";
        public const string StructureFileSuffix = ".structure.xml";
        public const string BookDifferencesFileSuffix = ".diff.xml";


        protected string ZefaniaXmlFilePath { get; set; }
        protected string ModuleShortName { get; set; }
        protected string ModuleDisplayName { get; set; }
        protected ModuleBibleStructure BibleStructure { get; set; }
        protected BibleBookDifferences BibleBookDifferences { get; set; }
        protected BibleBooksInfo BibleBooksInfo { get; set; }
        protected XMLBIBLE BibleContent { get; set; }

        public ZefaniaXmlConverterForm()
        {
            InitializeComponent();
        }

        private void ZefaniaXmlConverterForm_Load(object sender, EventArgs e)
        {
            EnableAll(false, this.Controls, btnZefaniaXmlFilePath);

            BindControls();            
        }

        private void EnableAll(bool enabled, Control.ControlCollection controls, params Control[] except)
        {   
            foreach (Control control in controls)
            {
                EnableAll(enabled, control.Controls, except);

                if (!except.Contains(control))
                    control.Enabled = enabled;
            }
        }      

        private void BindControls()
        {
        
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }        

        private void btnOk_Click(object sender, EventArgs e)
        {
            var converter = new ZefaniaXmlConverter("ibs", "Современный перевод (Всемирный Библейский Переводческий Центр)",
                @"C:\Users\lux_demko\Desktop\temp\Dropbox\Holy Bible\ForGenerating\ibs\bible.xml",
                Utils.LoadFromXmlString<BibleBooksInfo>(Properties.Resources.BibleBooskInfo_rst), @"c:\temp\ibsZefania", "ru",
                PredefinedNotebooksInfo.Russian, Utils.LoadFromXmlString<BibleTranslationDifferences>(Properties.Resources.BibleTranslationDifferences_rst),  // вот эти тоже часто надо менять                
                "{0} глава. {1}",
                PredefinedSectionsInfo.None, false, null, null,
                //PredefinedSectionsInfo.RSTStrong, true, "Стронга", 14700,   // параметры для стронга
                new Version(2, 0), true,
                ZefaniaXmlConverter.ReadParameters.None);  // и про эту не забыть

            converter.Convert();
        }        

        private void chkNotebookBibleGenerate_CheckedChanged(object sender, EventArgs e)
        {
            cbNotebookBible.Enabled = !((CheckBox)sender).Checked;
            tbNotebookBibleName.Enabled = ((CheckBox)sender).Checked;
        }

        private void chkNotebookBibleCommentsGenerate_CheckedChanged(object sender, EventArgs e)
        {
            cbNotebookBibleComments.Enabled = !((CheckBox)sender).Checked;
            tbNotebookBibleCommentsName.Enabled = ((CheckBox)sender).Checked;
        }

        private void chkNotebookSummaryOfNotesGenerate_CheckedChanged(object sender, EventArgs e)
        {
            cbNotebookSummaryOfNotes.Enabled = !((CheckBox)sender).Checked;
            tbNotebookSummaryOfNotesName.Enabled = ((CheckBox)sender).Checked;
        }


        private void cbNotebookBibleStudyUseFromFile_CheckedChanged(object sender, EventArgs e)
        {
            cbNotebookBibleStudy.Enabled = !((CheckBox)sender).Checked; 
            tbNotebookBibleStudyFilePath.Enabled = ((CheckBox)sender).Checked;
            btnNotebookBibleStudyFilePath.Enabled = ((CheckBox)sender).Checked;
        }   

        private void btnZefaniaXmlFilePath_Click(object sender, EventArgs e)
        {
            if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                this.ZefaniaXmlFilePath = openFileDialog.FileName;
                this.ModuleShortName = Path.GetFileNameWithoutExtension(this.ZefaniaXmlFilePath);
                LoadFiles();

                EnableAll(true, this.Controls);

                ChangeControlsStatus();                
            }
        }

        private void LoadFiles()
        {
            var moduleFolder = Path.GetDirectoryName(this.ZefaniaXmlFilePath);
            var localeFolder = Path.GetDirectoryName(moduleFolder);                        

            if (File.Exists(Path.Combine(moduleFolder, this.ModuleShortName + StructureFileSuffix))
        }

        private void ChangeControlsStatus()
        {
            tbZefaniaXmlFilePath.Text = this.ZefaniaXmlFilePath;
            tbVersion.Text = "2.0";
            tbLocale.Text = Path.GetFileName(Path.GetDirectoryName(openFileDialog.FileName)).ToLower();
        }       
    }
}
