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
using System.Diagnostics;
using Microsoft.Office.Interop.OneNote;
using System.Xml;
using System.Threading;

namespace BibleConfigurator
{
    public partial class ZefaniaXmlConverterForm : Form
    {
        private const string Const_StructureFileName = "structure.xml";
        private const string Const_BooksInfoFileName = "books.xml";        
        private const string Const_BookDifferencesFileSuffix = ".diff.xml";
        private const string Const_OutputDirectoryName = "Output";


        public bool NeedToCheckModule { get; set; }
        public string ConvertedModuleShortName { get; set; }

        protected Microsoft.Office.Interop.OneNote.Application OneNoteApp { get; set; }

        protected string ZefaniaXmlFilePath { get; set; }
        protected string ModuleShortName { get; set; }
        protected string ModuleDisplayName { get; set; }        
        protected string Locale { get; set; }
        protected bool BibleIsStrong { get; set; }
        protected string ChapterPageNameTemplate { get; set; }
        protected string NotebookBibleName { get; set; }
        protected string NotebookBibleCommentsName { get; set; }
        protected string NotebookSummaryOfNotesName { get; set; }

        protected string ModuleDirectory { get; set; }
        protected string LocaleDirectory { get; set; }
        protected string OutputLocaleDirectory { get; set; }
        protected string OutputModuleDirectory { get; set; }
        protected string RootDirectory { get; set; }
        protected string InputDirectory { get; set; }
        protected string MarkingWordsSectionFilePath { get; set; }        
        
        protected Dictionary<ContainerType, List<string>> ExistingNotebooks { get; set; }

        protected NotebooksStructure NotebooksStructure { get; set; }
        protected BibleTranslationDifferences BibleBookDifferences { get; set; }
        protected BibleBooksInfo BibleBooksInfo { get; set; }
        protected ModuleInfo ExistingOutputModule { get; set; }
        protected XMLBIBLE BibleContent { get; set; }


        private MainForm _mainForm;
        private LongProcessLogger _formLogger;
        private FileSystemWatcher _fileWatcher;
        private DateTime _startTime;
        private static object _locker = new object();
        private volatile Dictionary<string, string> _notebookFilesToWatch = new Dictionary<string, string>();
        private bool _allNotebooksPublished = false;

        public ZefaniaXmlConverterForm(Microsoft.Office.Interop.OneNote.Application oneNoteApp, MainForm mainForm)
        {
            InitializeComponent();
            OneNoteApp = oneNoteApp;
            _mainForm = mainForm;
            _formLogger = new LongProcessLogger(_mainForm);
        }

        private void ZefaniaXmlConverterForm_Load(object sender, EventArgs e)
        {
            FormExtensions.EnableAll(false, Controls, tbZefaniaXmlFilePath, btnZefaniaXmlFilePath, btnClose);
            this.Top = this.Top - 15;
        }              

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }        

        private void btnOk_Click(object sender, EventArgs e)
        {
            try
            {
                BibleCommon.Services.Logger.Init("ZefaniaXmlConverterForm");

                _startTime = DateTime.Now;
                BibleCommon.Services.Logger.LogMessage("{0}: {1}", BibleCommon.Resources.Constants.StartTime, _startTime.ToLongTimeString());


                var initMessage = "Start converting";
                _mainForm.PrepareForLongProcessing(5 + (chkNotebookBibleGenerate.Checked
                                                        ? ModulesManager.GetBibleChaptersCount(BibleContent, true)
                                                        : 0),
                                                    1, initMessage);

                _formLogger.LogMessage(initMessage);
                BibleCommon.Services.Logger.LogMessage(initMessage);
                FormExtensions.EnableAll(false, this.Controls, btnClose);
                System.Windows.Forms.Application.DoEvents();

                

                var converter = new ZefaniaXmlConverter(tbShortName.Text, tbDisplayName.Text, BibleContent, BibleBooksInfo,
                                                            tbResultDirectory.Text, tbLocale.Text, NotebooksStructure,
                                                            BibleBookDifferences, ChapterPageNameTemplate,
                                                            BibleIsStrong && !chkRemoveStrongNumbers.Checked,
                                                            new Version(tbVersion.Text),
                                                            chkNotebookBibleGenerate.Checked,
                                                            _formLogger,
                                                            chkRemoveStrongNumbers.Checked ? ZefaniaXmlConverter.ReadParameters.RemoveStrongs : ZefaniaXmlConverter.ReadParameters.None);

                if (chkNotebookBibleGenerate.Checked)
                    _formLogger.Preffix = "Creating the Bible: ";                

                converter.Convert();

                _formLogger.Preffix = string.Empty;
                _formLogger.LogMessage("Publishing notebooks");

                RemovePrevModuleFile();

                CopyStrongSections();

                PublishNotebooks(converter);

                if (!NeedToWaitFileWatcher)
                {
                    CreateZipFileAndFinish();
                    CloseResources();
                }
                else
                    _formLogger.LogMessage("Saving Notebooks files");
            }
            catch (ProcessAbortedByUserException)
            {
                BibleCommon.Services.Logger.LogMessage(BibleCommon.Resources.Constants.ProcessAbortedByUser);
                _mainForm.LongProcessingDone(BibleCommon.Resources.Constants.ProcessAbortedByUser);

                CloseResources();
            }
            catch (Exception ex)
            {
                FormLogger.LogError(ex);
            }            
        }

        private void CreateZipFileAndFinish()
        {
            _formLogger.LogMessage("Creating module zip file");
            var resultFilePath = CreateZipFile();

            //FormLogger.LogMessage("Finished");

            ConvertedModuleShortName = ModuleShortName;

            string finalMessage = "Module was created";
            _mainForm.LongProcessingDone(finalMessage);
            BibleCommon.Services.Logger.LogMessage(finalMessage);

            BibleCommon.Services.Logger.LogMessage(" {0}", DateTime.Now.Subtract(_startTime));
            BibleCommon.Services.Logger.Done();       

            if (chkCheckModule.Checked)
            {
                bool moduleWasAdded;
                bool needToReload = _mainForm.AddNewModule(resultFilePath, out moduleWasAdded);
                if (needToReload)
                    _mainForm.ReLoadParameters(false);

                if (moduleWasAdded)
                    NeedToCheckModule = true;
            }
            else
                Process.Start("explorer.exe", "/select," + resultFilePath);

            this.Invoke(Close);
        }

        private void RemovePrevModuleFile()
        {
            var files = Directory.GetFiles(tbResultDirectory.Text, "*" + BibleCommon.Consts.Constants.FileExtensionIsbt);
            foreach (var file in files)
            {
                File.Delete(file);
            }
        }

        private string CreateZipFile()
        {
            var resultFilePath = Path.Combine(tbResultDirectory.Text, ModuleShortName + BibleCommon.Consts.Constants.FileExtensionIsbt);
            ZipLibHelper.PackfilesToZip(tbResultDirectory.Text, resultFilePath);

            return resultFilePath;
        }

        private void CopyStrongSections()
        {
            if (BibleIsStrong && !chkRemoveStrongNumbers.Checked)
            {
                foreach (var sectionFile in Directory.GetFiles(ModuleDirectory, "*.one"))
                {
                    File.Copy(sectionFile, Path.Combine(tbResultDirectory.Text, Path.GetFileName(sectionFile)), true);
                }
            }
        }

        private void PublishNotebooks(ZefaniaXmlConverter converter)
        {
            if (chkNotebookBibleGenerate.Checked)
            {
                AddMarkingWordsSection(converter.BibleNotebookId);
                PublishNotebook(converter.BibleNotebookId, Path.Combine(tbResultDirectory.Text, NotebookBibleName + BibleCommon.Consts.Constants.FileExtensionOnepkg), true);
            }
            else
            {
                CopyExistingNotebookFile(cbNotebookBible);
            }

            CopyExistingNotebookFile(cbNotebookBibleStudy);

            if (chkNotebookBibleCommentsGenerate.Checked)
            {
                var notebookId = NotebookGenerator.GenerateBibleCommentsNotebook(OneNoteApp, NotebookBibleCommentsName, converter.ModuleInfo.BibleStructure, NotebooksStructure, false);
                PublishNotebook(notebookId, Path.Combine(tbResultDirectory.Text, NotebookBibleCommentsName + BibleCommon.Consts.Constants.FileExtensionOnepkg), true);                
            }
            else
            {
                CopyExistingNotebookFile(cbNotebookBibleComments);
            }

            if (chkNotebookSummaryOfNotesGenerate.Checked)
            {
                var notebookId = NotebookGenerator.GenerateBibleCommentsNotebook(OneNoteApp, NotebookSummaryOfNotesName, converter.ModuleInfo.BibleStructure, NotebooksStructure, true);
                PublishNotebook(notebookId, Path.Combine(tbResultDirectory.Text, NotebookSummaryOfNotesName + BibleCommon.Consts.Constants.FileExtensionOnepkg), true);                
            }
            else
            {
                CopyExistingNotebookFile(cbNotebookSummaryOfNotes);
            }           

            _allNotebooksPublished = true;
        }

        private void AddFileForWatching(string notebookFilePath, string notebookId)
        {
            if (_fileWatcher == null)
            {
                _fileWatcher = new FileSystemWatcher(tbResultDirectory.Text, "*" + BibleCommon.Consts.Constants.FileExtensionOnepkg);

                _fileWatcher.Changed += new FileSystemEventHandler(_fileWatcher_Changed);
                _fileWatcher.EnableRaisingEvents = true;
            }

            lock (_locker)
            {
                _notebookFilesToWatch.Add(notebookFilePath, notebookId);
            }
        }

        private bool NeedToWaitFileWatcher
        {
            get
            {
                return _notebookFilesToWatch.Count > 0 || !_allNotebooksPublished;
            }
        }

        private void _fileWatcher_Changed(object sender, FileSystemEventArgs e)
        {
            try
            {
                if (_mainForm.StopExternalProcess)
                    throw new ProcessAbortedByUserException();

                if (new FileInfo(e.FullPath).Length > 0)
                {
                    if (_notebookFilesToWatch.ContainsKey(e.FullPath))
                    {
                        lock (_locker)
                        {
                            if (_notebookFilesToWatch.ContainsKey(e.FullPath))
                            {
                                var notebookId = _notebookFilesToWatch[e.FullPath];                                
                                if (!string.IsNullOrEmpty(notebookId))                                    
                                    OneNoteApp.CloseNotebook(notebookId);
                                _notebookFilesToWatch.Remove(e.FullPath);                                

                                if (!NeedToWaitFileWatcher)
                                {
                                    Thread.Sleep(1000);
                                    CreateZipFileAndFinish();
                                    CloseResources();
                                }                                
                            }
                        }
                    }
                }
            }
            catch (ProcessAbortedByUserException)
            {
                BibleCommon.Services.Logger.LogMessage("Process aborted by user");

                CloseResources();
            }
        }

        private void CloseResources()
        {
            if (_fileWatcher != null)
            {
                _fileWatcher.EnableRaisingEvents = false;
                _fileWatcher.Dispose();
            }                              
        }


        private void AddMarkingWordsSection(string notebookId)
        {
            if (string.IsNullOrEmpty(MarkingWordsSectionFilePath))
                throw new Exception("Marking section file was not found");

            XmlNamespaceManager xnm;
            var notebookDoc = OneNoteUtils.GetHierarchyElement(OneNoteApp, notebookId, HierarchyScope.hsSelf, out xnm);
            var notebookPath = (string)notebookDoc.Root.Attribute("path");
            var markingWordsSectionName = Path.GetFileName(MarkingWordsSectionFilePath);

            File.Copy(MarkingWordsSectionFilePath, Path.Combine(notebookPath, markingWordsSectionName), true);
            string markingWordsSectionId;
            OneNoteApp.OpenHierarchy(markingWordsSectionName, notebookId, out markingWordsSectionId, CreateFileType.cftSection);
        }

        private void PublishNotebook(string notebookId, string targetFilePath, bool closeNotebook)
        {
            if (File.Exists(targetFilePath))
                File.Delete(targetFilePath);


            AddFileForWatching(targetFilePath, closeNotebook ? notebookId : null);
            OneNoteApp.Publish(notebookId, targetFilePath, PublishFormat.pfOneNotePackage);                        
        }

        private void CopyExistingNotebookFile(ComboBox cb)
        {
            var selectedFilePath = (string)cb.SelectedItem;
            if (!string.IsNullOrEmpty(selectedFilePath))
            {
                if (selectedFilePath.StartsWith("\\"))
                    selectedFilePath = selectedFilePath.Substring(1);

                var sourceFilePath = Path.Combine(RootDirectory, selectedFilePath);
                var targetFilePath = Path.Combine(tbResultDirectory.Text, Path.GetFileName(selectedFilePath));

                if (!sourceFilePath.Equals(targetFilePath, StringComparison.InvariantCultureIgnoreCase))
                    File.Copy(sourceFilePath, targetFilePath, true);
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
        }

        private void chkNotebookBibleCommentsGenerate_CheckedChanged(object sender, EventArgs e)
        {
            cbNotebookBibleComments.Enabled = !((CheckBox)sender).Checked;            
        }

        private void chkNotebookSummaryOfNotesGenerate_CheckedChanged(object sender, EventArgs e)
        {
            cbNotebookSummaryOfNotes.Enabled = !((CheckBox)sender).Checked;            
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

                    FormExtensions.EnableAll(true, Controls);

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
            InputDirectory = Path.GetDirectoryName(LocaleDirectory);
            RootDirectory = Path.GetDirectoryName(InputDirectory);

            Locale = Path.GetFileName(LocaleDirectory);

            OutputLocaleDirectory = Path.Combine(Path.Combine(Path.GetDirectoryName(Path.GetDirectoryName(LocaleDirectory)), Const_OutputDirectoryName), Locale);           
            if (!Directory.Exists(OutputLocaleDirectory))
                Directory.CreateDirectory(OutputLocaleDirectory);

            OutputModuleDirectory = Path.Combine(OutputLocaleDirectory, ModuleShortName);
            if (!Directory.Exists(OutputModuleDirectory))
                Directory.CreateDirectory(OutputModuleDirectory);           
        }

        private void LoadFiles()
        {
            var structureFilePath = GetExistingFile(
                                        Path.Combine(ModuleDirectory, string.Format("{0}.{1}", ModuleShortName, Const_StructureFileName)),
                                        Path.Combine(LocaleDirectory, Const_StructureFileName),
                                        Path.Combine(InputDirectory, Const_StructureFileName));
            var booksFilePath = GetExistingFile(
                                        Path.Combine(ModuleDirectory, string.Format("{0}.{1}", ModuleShortName, Const_BooksInfoFileName)),
                                        Path.Combine(LocaleDirectory, Const_BooksInfoFileName),
                                        Path.Combine(InputDirectory, Const_BooksInfoFileName));
            var diffFilePath = Path.Combine(ModuleDirectory, ModuleShortName + Const_BookDifferencesFileSuffix);            
            var existingModuleFilePath = Path.Combine(OutputModuleDirectory, BibleCommon.Consts.Constants.ManifestFileName);

            var invalidXmlFiles = Directory.GetFiles(ModuleDirectory, string.Format("*.xml", ModuleShortName))
                                    .Where(f =>
                                        !Path.GetFileName(f).ToLower().StartsWith(ModuleShortName));
            if (invalidXmlFiles.Count() > 0)
                throw new Exception(string.Format("There are invalid .xml files in '{0}': {1}", ModuleDirectory,
                    string.Join(Environment.NewLine, invalidXmlFiles.Select(f => Path.GetFileName(f)).ToArray())));
            

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
            LoadMarkingWordsSectionFilePath();
        }
        
        private void LoadMarkingWordsSectionFilePath()
        {
            MarkingWordsSectionFilePath = LoadMarkingWordsSectionFilePath(LocaleDirectory);

            if (string.IsNullOrEmpty(MarkingWordsSectionFilePath))
                MarkingWordsSectionFilePath = LoadMarkingWordsSectionFilePath(InputDirectory);
        }

        private static string LoadMarkingWordsSectionFilePath(string directory)
        {
            var files = Directory.GetFiles(directory, "*.one").Where(file => file.EndsWith(".one")).ToArray();
            if (files.Length > 0)
                return files[0];

            return null;
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

            if (string.IsNullOrEmpty(BibleBooksInfo.ChapterPageNameTemplate))
                throw new Exception("BibleBooksInfo.ChapterPageNameTemplate is null");

            ChapterPageNameTemplate = BibleBooksInfo.ChapterPageNameTemplate;            
        }        

        private void ChangeControlsState()
        {
            tbZefaniaXmlFilePath.Text = ZefaniaXmlFilePath;
            tbZefaniaXmlFilePath.ReadOnly = true;

            tbShortName.Text = ExistingOutputModule != null ? ExistingOutputModule.ShortName : ModuleShortName;
            tbVersion.Text = ExistingOutputModule != null ? ExistingOutputModule.Version.ToString() : "2.0";
            tbLocale.Text = ExistingOutputModule != null ? ExistingOutputModule.Locale : Locale;

            tbDisplayName.Text = GetModuleDisplayName(BibleContent);
            if (string.IsNullOrEmpty(tbDisplayName.Text))
                tbDisplayName.Text = ExistingOutputModule != null ? ExistingOutputModule.DisplayName : ModuleShortName;

            tbResultDirectory.Text = OutputModuleDirectory;
            folderBrowserDialog.SelectedPath = OutputModuleDirectory;

            NotebookBibleName = SetNotebookParams(cbNotebookBible, chkNotebookBibleGenerate, ContainerType.Bible, !BibleIsStrong);
            SetNotebookParams(cbNotebookBibleStudy, null, ContainerType.BibleStudy, null);
            NotebookBibleCommentsName = SetNotebookParams(cbNotebookBibleComments, chkNotebookBibleCommentsGenerate, ContainerType.BibleComments, true);
            NotebookSummaryOfNotesName = SetNotebookParams(cbNotebookSummaryOfNotes, chkNotebookSummaryOfNotesGenerate, ContainerType.BibleNotesPages, true);

            if (!BibleIsStrong)
                chkRemoveStrongNumbers.Enabled = false;
        }

        private string SetNotebookParams(ComboBox cb, CheckBox chk, ContainerType notebookType, bool? generateByDefault)
        {
            string result = null;
            var notebookInfo = NotebooksStructure.Notebooks.FirstOrDefault(n => n.Type == notebookType);
            if (notebookInfo != null)
            {
                cb.DataSource = GetExistingNotebooks(ExistingNotebooks, notebookType, notebookInfo.Name);
                SetCbNotebooksValue(cb, chk, notebookType, generateByDefault);

                result = Path.GetFileNameWithoutExtension(notebookInfo.Name);
            }
            else
            {   
                if (chk != null)
                    chk.Enabled = false;
                cb.Enabled = false;
            }

            return result;
        }

        private static bool IsStrong(XMLBIBLE bibleContent)
        {
            return bibleContent.Books.Any(b =>
                                    b.Chapters.Any(c =>
                                        c.Verses.Any(v =>
                                            v.Items != null
                                            && v.Items.Any(item =>
                                                item is GRAM || item is gr))));
        }

        private static void SetCbNotebooksValue(ComboBox cb, CheckBox chk, ContainerType notebookType, bool? generateByDefault)
        {
            if (cb.Items.Count > 0 || !generateByDefault.GetValueOrDefault(true))
            {
                if (cb.Items.Count > 0)
                    cb.SelectedIndex = 0;                
            }
            else if (chk != null)                
                chk.Checked = true;
            else
                throw new Exception(string.Format("Notebook file for {0} was not found.", notebookType));
        }

        private static List<string> GetExistingNotebooks(Dictionary<ContainerType, List<string>> dictionary, ContainerType key, 
                                            string notebookName)
        {
            if (dictionary.ContainsKey(key))            
                return dictionary[key].Where(n => Path.GetFileName(n).Equals(notebookName, StringComparison.InvariantCultureIgnoreCase)).ToList();
            else 
                return new List<string>();
        }

        private void LoadExistingNotebooks()
        {
            ExistingNotebooks = new Dictionary<ContainerType, List<string>>();

            if (ExistingOutputModule != null)            
                LoadNotebookFiles(OutputModuleDirectory, 4, null, true);            

            LoadNotebookFiles(ModuleDirectory, 4, ContainerType.BibleStudy, false);
            LoadNotebookFiles(LocaleDirectory, 3, ContainerType.BibleStudy, false);
            LoadNotebookFiles(InputDirectory, 2, ContainerType.BibleStudy, false);
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

        private void ZefaniaXmlConverterForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            _formLogger.AbortedByUsers = true;
        }

        private void ZefaniaXmlConverterForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            OneNoteApp = null;
            _mainForm = null;
            _formLogger.Dispose();
        }
    }
}
