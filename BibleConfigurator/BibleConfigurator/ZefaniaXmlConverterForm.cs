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
using BibleCommon.Services;

namespace BibleConfigurator
{
    public partial class ZefaniaXmlConverterForm : Form
    {
        public const string Const_LocaleStructureFilePath = "structure.xml";
        public const string Const_LocaleBooksInfoFilePath = "books.xml";
        public const string Const_StructureFileSuffix = ".structure.xml";
        public const string Const_BookDifferencesFileSuffix = ".diff.xml";
        public const string Const_OutputDirectoryName = "Output";


        protected Microsoft.Office.Interop.OneNote.Application OneNoteApp { get; set; }

        protected string ZefaniaXmlFilePath { get; set; }
        protected string ModuleShortName { get; set; }
        protected string ModuleDisplayName { get; set; }        
        protected string Locale { get; set; }
        protected bool BibleIsStrong { get; set; }
        protected string ChapterSectionNameTemplate { get; set; }

        protected string ModuleDirectory { get; set; }
        protected string LocaleDirectory { get; set; }
        protected string OutputLocaleDirectory { get; set; }
        protected string OutputModuleDirectory { get; set; }
        protected string RootDirectory { get; set; }
        
        protected Dictionary<ContainerType, List<string>> ExistingNotebooks { get; set; }

        protected NotebooksStructure NotebooksStructure { get; set; }
        protected BibleTranslationDifferences BibleBookDifferences { get; set; }
        protected BibleBooksInfo BibleBooksInfo { get; set; }
        protected ModuleInfo ExistingOutputModule { get; set; }
        protected XMLBIBLE BibleContent { get; set; }

        public ZefaniaXmlConverterForm()
        {
            InitializeComponent();            
        }

        private void ZefaniaXmlConverterForm_Load(object sender, EventArgs e)
        {
            EnableAll(false, Controls, tbZefaniaXmlFilePath, btnZefaniaXmlFilePath, btnClose);

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
            Close();
        }        

        private void btnOk_Click(object sender, EventArgs e)
        {
            try
            {
                var converter = new ZefaniaXmlConverter(tbShortName.Text, tbDisplayName.Text, BibleContent, BibleBooksInfo,
                                                            tbResultDirectory.Text, tbLocale.Text, NotebooksStructure,
                                                            BibleBookDifferences, ChapterSectionNameTemplate,
                                                            BibleIsStrong && !chkRemoveStrongNumbers.Checked,
                                                            new Version(tbVersion.Text),
                                                            chkNotebookBibleGenerate.Checked,
                                                            chkRemoveStrongNumbers.Checked ? ZefaniaXmlConverter.ReadParameters.RemoveStrongs : ZefaniaXmlConverter.ReadParameters.None);

                converter.Convert();

                OneNoteApp = new Microsoft.Office.Interop.OneNote.Application();

                if (chkNotebookBibleGenerate.Checked)
                {
                    OneNoteApp.Publish(converter.NotebookId, Path.Combine(tbResultDirectory.Text, tbNotebookBibleName.Text + BibleCommon.Consts.Constants.FileExtensionOnepkg));
                }
                else
                {
                    CopyExistingNotebookFile(cbNotebookBible);
                }

                CopyExistingNotebookFile(cbNotebookBibleStudy);                

                if (chkNotebookBibleCommentsGenerate.Checked)
                {
                    var notebookId = NotebookGenerator.GenerateBibleCommentsNotebook(OneNoteApp, tbNotebookBibleCommentsName.Text, converter.ModuleInfo.BibleStructure, NotebooksStructure, false);
                    OneNoteApp.Publish(notebookId, Path.Combine(tbResultDirectory.Text, tbNotebookBibleCommentsName.Text + BibleCommon.Consts.Constants.FileExtensionOnepkg));
                }
                else
                {
                    CopyExistingNotebookFile(cbNotebookBibleComments);
                }

                if (chkNotebookSummaryOfNotesGenerate.Checked)
                {
                    var notebookId = NotebookGenerator.GenerateBibleCommentsNotebook(OneNoteApp, tbNotebookSummaryOfNotesName.Text, converter.ModuleInfo.BibleStructure, NotebooksStructure, true);
                    OneNoteApp.Publish(notebookId, Path.Combine(tbResultDirectory.Text, tbNotebookSummaryOfNotesName.Text + BibleCommon.Consts.Constants.FileExtensionOnepkg));
                }
                else
                {
                    CopyExistingNotebookFile(cbNotebookSummaryOfNotes);
                }

                foreach(var sectionFile in Directory.GetFiles(ModuleDirectory, "*.one"))
                {
                    File.Copy(sectionFile, Path.Combine(tbResultDirectory.Text, Path.GetFileName(sectionFile)), true);
                }
            }
            catch (Exception ex)
            {
                FormLogger.LogError(ex);
            }
            finally
            {
                OneNoteApp = null;
            }
        }

        private void CopyExistingNotebookFile(ComboBox cb)
        {
            if (cb.Enabled && !string.IsNullOrEmpty(cb.SelectedText))
            {
                File.Copy(Path.Combine(RootDirectory, cb.SelectedText), Path.Combine(tbResultDirectory.Text, Path.GetFileName(cb.SelectedText)), true);
            }
        }

        private void CheckAndCorrectSections()
        {
            var sectionFiles = Directory.GetFiles(ModuleDirectory, "*.one");
            foreach (var sectionInfo in NotebooksStructure.Sections)
            {
                if (!sectionFiles.Any(sf => Path.GetFileName(sf).Equals(sectionInfo.Name, StringComparison.InvariantCultureIgnoreCase)))
                    throw new Exception(string.Format("File '{0}' was not found", Path.Combine(ModuleDirectory, sectionInfo.Name)));
            }
            foreach (var sectionFile in sectionFiles)
            {
                var sectionFileName = Path.GetFileName(sectionFile);
                if (!NotebooksStructure.Sections.Any(s => s.Name.Equals(sectionFileName, StringComparison.InvariantCultureIgnoreCase)))
                    NotebooksStructure.Sections.Add(new SectionInfo() { Name = sectionFileName });
            }
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

        private void btnZefaniaXmlFilePath_Click(object sender, EventArgs e)
        {
            if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    ZefaniaXmlFilePath = openFileDialog.FileName;

                    LoadBaseInfo();

                    LoadFiles();

                    LoadAdditionalInfo();

                    EnableAll(true, Controls);

                    ChangeControlsState();                    
                }
                catch (Exception ex)
                {
                    FormLogger.LogError(ex);
                }
            }
        }        

        private void LoadBaseInfo()
        {
            ModuleShortName = Path.GetFileNameWithoutExtension(ZefaniaXmlFilePath).ToLower();            

            ModuleDirectory = Path.GetDirectoryName(ZefaniaXmlFilePath);
            LocaleDirectory = Path.GetDirectoryName(ModuleDirectory);            
            RootDirectory = Path.GetDirectoryName(LocaleDirectory);

            Locale = Path.GetFileName(LocaleDirectory);

            OutputLocaleDirectory = Path.Combine(Path.Combine(Path.GetDirectoryName(Path.GetDirectoryName(LocaleDirectory)), Const_OutputDirectoryName), Locale);           
            if (!Directory.Exists(OutputLocaleDirectory))
                Directory.CreateDirectory(OutputLocaleDirectory);

            OutputModuleDirectory = Path.Combine(OutputLocaleDirectory, ModuleShortName);
            if (!Directory.Exists(OutputModuleDirectory))
                Directory.CreateDirectory(OutputModuleDirectory);

            var invalidXmlFiles = Directory.GetFiles(ModuleDirectory, string.Format("*.xml", ModuleShortName))
                                                .Where(f => 
                                                    !Path.GetFileName(f).ToLower().StartsWith(ModuleShortName));

            if (invalidXmlFiles.Count() > 0)
                throw new Exception(string.Format("There are invalid .xml files in '{0}': {1}", ModuleDirectory, 
                    string.Join(Environment.NewLine, invalidXmlFiles.Select(f => Path.GetFileName(f)).ToArray())));
        }

        private void LoadFiles()
        {
            var structureFilePath = GetExistingFile(
                                        Path.Combine(ModuleDirectory, ModuleShortName + Const_StructureFileSuffix),
                                        Path.Combine(LocaleDirectory, Const_LocaleStructureFilePath));
            var booksFilePath = GetExistingFile(Path.Combine(LocaleDirectory, Const_LocaleBooksInfoFilePath));
            var diffFilePath = Path.Combine(ModuleDirectory, ModuleShortName + Const_BookDifferencesFileSuffix);            
            var existingModuleFilePath = Path.Combine(OutputModuleDirectory, BibleCommon.Consts.Constants.ManifestFileName);
            

            BibleContent = Utils.LoadFromXmlFile<XMLBIBLE>(ZefaniaXmlFilePath);
            NotebooksStructure = Utils.LoadFromXmlFile<NotebooksStructure>(structureFilePath);
            BibleBooksInfo = Utils.LoadFromXmlFile<BibleBooksInfo>(booksFilePath);

            BibleBookDifferences = File.Exists(diffFilePath)
                                            ? Utils.LoadFromXmlFile<BibleTranslationDifferences>(diffFilePath)
                                            : new BibleTranslationDifferences();

            ExistingOutputModule = File.Exists(existingModuleFilePath)
                                            ? Utils.LoadFromXmlFile<ModuleInfo>(existingModuleFilePath)
                                            : null;            

            LoadExistingNotebooks();
            CheckAndCorrectSections();            
        }

        private static string GetExistingFile(params string[] filesPath)
        {
            foreach (var file in filesPath)
            {
                if (File.Exists(file))
                    return file;
            }

            throw new Exception("No one file exists: " + Environment.NewLine + string.Join(Environment.NewLine, filesPath));
        }

        private void LoadAdditionalInfo()
        {
            BibleIsStrong = IsStrong(BibleContent);

            if (string.IsNullOrEmpty(BibleBooksInfo.ChapterString))
                throw new Exception("BibleBooksInfo.ChapterString is null");

            ChapterSectionNameTemplate = string.Format("{{0}} {0}. {{1}}", BibleBooksInfo.ChapterString.ToLower());
        }

        private void ChangeControlsState()
        {
            tbZefaniaXmlFilePath.Text = ZefaniaXmlFilePath;
            tbZefaniaXmlFilePath.ReadOnly = true;

            tbShortName.Text = ModuleShortName;
            tbVersion.Text = "2.0";
            tbLocale.Text = Locale;
            tbDisplayName.Text = GetModuleDisplayName(BibleContent);
            if (string.IsNullOrEmpty(tbDisplayName.Text))
                tbDisplayName.Text = ModuleShortName;
            tbResultDirectory.Text = OutputModuleDirectory;
            folderBrowserDialog.SelectedPath = OutputModuleDirectory;

            SetNotebookParams(cbNotebookBible, chkNotebookBibleGenerate, tbNotebookBibleName, ContainerType.Bible, !BibleIsStrong);
            SetNotebookParams(cbNotebookBibleStudy, null, null, ContainerType.BibleStudy, null);
            SetNotebookParams(cbNotebookBibleComments, chkNotebookBibleCommentsGenerate, tbNotebookBibleCommentsName, ContainerType.BibleComments, true);
            SetNotebookParams(cbNotebookSummaryOfNotes, chkNotebookSummaryOfNotesGenerate, tbNotebookSummaryOfNotesName, ContainerType.BibleNotesPages, true);

            if (!BibleIsStrong)
                chkRemoveStrongNumbers.Enabled = false;
        }

        private void SetNotebookParams(ComboBox cb, CheckBox chk, TextBox tb, ContainerType notebookType, bool? generateByDefault)
        {
            var notebookInfo = NotebooksStructure.Notebooks.FirstOrDefault(n => n.Type == notebookType);
            if (notebookInfo != null)
            {
                cb.DataSource = GetDictionaryValueOrDefault<ContainerType, List<string>>(ExistingNotebooks, notebookType, new List<string>());
                SetCbNotebooksValue(cb, chk, tb, notebookType, generateByDefault);
                if (tb != null)
                    tb.Text = Path.GetFileNameWithoutExtension(notebookInfo.Name);
            }
            else
            {
                if (tb != null)
                    tb.Enabled = false;
                if (chk != null)
                    chk.Enabled = false;
                cb.Enabled = false;
            }
        }

        private static bool IsStrong(XMLBIBLE bibleContent)
        {
            return bibleContent.Books.Any(b =>
                                    b.Chapters.Any(c =>
                                        c.Verses.Any(v =>
                                            v.Items.Any(item =>
                                                item is GRAM || item is gr))));
        }

        private void SetCbNotebooksValue(ComboBox cb, CheckBox chk, TextBox tb, ContainerType notebookType, bool? generateByDefault)
        {
            if (cb.Items.Count > 0 || !generateByDefault.GetValueOrDefault(true))
            {
                if (cb.Items.Count > 0)
                    cb.SelectedIndex = 0;
                if (tb != null)
                    tb.Enabled = false;
            }
            else if (chk != null)                
                chk.Checked = true;
            else
                throw new Exception(string.Format("Notebook file for {0} was not found.", notebookType));
        }

        private static TValue GetDictionaryValueOrDefault<TKey, TValue>(Dictionary<TKey, TValue> dictionary, TKey key, TValue defaultValue)
        {
            if (dictionary.ContainsKey(key))
                return dictionary[key];
            else 
                return defaultValue;
        }

        private void LoadExistingNotebooks()
        {
            ExistingNotebooks = new Dictionary<ContainerType, List<string>>();

            if (ExistingOutputModule != null)            
                LoadNotebookFiles(OutputModuleDirectory, 3, null, true);            

            LoadNotebookFiles(ModuleDirectory, 3, ContainerType.BibleStudy, false);
            LoadNotebookFiles(LocaleDirectory, 3, ContainerType.BibleStudy, false);
        }

        private void LoadNotebookFiles(string directory, int level, ContainerType? notebookType, bool loadFromExistingModule)
        {
            if (Directory.Exists(directory))
            {
                var notebooks = Directory.GetFiles(directory, "*" + BibleCommon.Consts.Constants.FileExtensionOnepkg);
                foreach (var notebookFulPath in notebooks)
                {
                    var parts = notebookFulPath.Split(new char[] { '\\' });
                    var notebookFile = string.Empty;
                    for (int i = 1; i <= level; i++)
                        notebookFile = "\\" + parts[parts.Length - i] + notebookFile;

                    var key = notebookType.HasValue
                                            ? notebookType.Value
                                            : ExistingOutputModule.NotebooksStructure.Notebooks.First(n => 
                                                                        n.Name.Equals(parts[parts.Length - 1], StringComparison.InvariantCultureIgnoreCase)).Type;

                    if (!ExistingNotebooks.ContainsKey(key))
                        ExistingNotebooks.Add(key, new List<string>());

                    ExistingNotebooks[key].Add(notebookFile);
                }
            }
        }

        private static string GetModuleDisplayName(XMLBIBLE bible)
        {
            if (bible.INFORMATION != null)
            {
                for (var i = 0; i < bible.INFORMATION.ItemsElementName.Length; i++)
                {
                    if (bible.INFORMATION.ItemsElementName[i] == ItemsChoiceType.title)
                        return (string)bible.INFORMATION.Items[i];
                }
            }

            return null;
        }

        private void tbZefaniaXmlFilePath_MouseClick(object sender, MouseEventArgs e)
        {
            if (tbZefaniaXmlFilePath.Enabled && tbZefaniaXmlFilePath.Text == string.Empty)
                btnZefaniaXmlFilePath_Click(btnZefaniaXmlFilePath, null);
        }

        private void btnResultFilePath_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                tbResultDirectory.Text = folderBrowserDialog.SelectedPath;
        }      
    }
}
