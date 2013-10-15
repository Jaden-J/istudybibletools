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
using BibleConfigurator.Tools;
using BibleCommon.Consts;
using System.Runtime.InteropServices;
using BibleCommon.Common;
using BibleConfigurator.ModuleConverter;
using System.Globalization;
using BibleCommon.UI.Forms;
using BibleCommon.Handlers;

namespace BibleConfigurator
{
    public partial class MainForm : Form
    {
        internal class ComboBoxItem
        {
            public string Value { get; set; }
            public object Key { get; set; }

            public ComboBoxItem(object key)
            {
                this.Key = key;                
            }

            public ComboBoxItem(object key, string value)
                : this(key)
            {                
                this.Value = value;
            }

            public override string ToString()
            {
                return Value;
            }

            public override bool Equals(object obj)
            {
                if (obj == null)
                    return false;

                if (obj is string)
                    return this.Key.ToString() == (string)obj;

                if (!(obj is ComboBoxItem))
                    return false;                

                return this.Key == ((ComboBoxItem)obj).Key;                    
            }

            public override int GetHashCode()
            {
                return this.Key.GetHashCode();
            }
        }

        private Microsoft.Office.Interop.OneNote.Application _oneNoteApp = OneNoteUtils.CreateOneNoteAppSafe();

        private string SingleNotebookFromTemplatePath { get; set; }
        private string BibleNotebookFromTemplatePath { get; set; }
        private string BibleCommentsNotebookFromTemplatePath { get; set; }
        private string BibleNotesPagesNotebookFromTemplatePath { get; set; }
        private string BibleStudyNotebookFromTemplatePath { get; set; }

        private bool _wasSearchedSectionGroupsInSingleNotebook = false;       
        

        private const int LoadParametersAttemptsCount = 80;         // количество попыток загрузки параметров после команды создания записных книжек из шаблона
        private const int LoadParametersPauseBetweenAttempts = 5000;             // количество милисекунд ожидания между попытками загрузки параметров
        private const string LoadParametersImageFileName = "loader.gif";

        protected LongProcessLogger LongProcessLogger { get; set; }

        private NotebookParametersForm _notebookParametersForm = null;

        private bool _moduleWasChanged = false;
        private string _originalModuleShortName; // модуль, который изначально является текущим в системе
        
        public bool ShowModulesTabAtStartUp { get; set; }
        public bool NeedToSaveChangesAfterLoadingModuleAtStartUp { get; set; }
        public bool ToIndexBible { get; set; }
        public bool NotAskToIndexBible { get; set; }
        public bool CommitChangesAfterLoad { get; set; }
        public string ForceIndexDictionaryModuleName { get; set; }

        public bool IsModerator { get; set; }

        private Dictionary<string, string> _notebooks;
        private bool _needToRefreshCache = false;   // чтобы понимать, нужно ли обновлять кэш

        public MainForm(params string[] args)
        {
            this.SetFormUICulture();

            InitializeComponent();
            BibleCommon.Services.Logger.Init("BibleConfigurator");
            LongProcessLogger = new LongProcessLogger(this);

            IsModerator = args.Contains(Consts.ModeratorMode);
        }

        public bool StopLongProcess { get; set; }        

        private void btnOK_Click(object sender, EventArgs e)
        {
            try
            {
                CommitChanges(true);
            }
            catch (Exception ex)
            {                
                FormLogger.LogError(ex);
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnApply_Click(object sender, EventArgs e)
        {
            try
            {
                CommitChanges(false);
            }
            catch (Exception ex)
            {                
                FormLogger.LogError(ex);
            }
        }    

        internal void CommitChanges(bool closeForm)
        {
            ModuleInfo module = null;

            try
            {
                module = ModulesManager.GetCurrentModuleInfo();
            }
            catch (InvalidModuleException ex)
            {
                FormLogger.LogMessage(ex.Message);
                return;
            }           

            btnOK.Enabled = false;            
            btnApply.Enabled = false;
            bool lblWarningVisibilityBefore = lblWarning.Visible;
            lblWarning.Visible = false;
            this.TopMost = true;

            try
            {
                FormLogger.Initialize();

                if ((chkCreateBibleNotebookFromTemplate.Enabled && chkCreateBibleNotebookFromTemplate.Checked)
                    || (chkCreateBibleStudyNotebookFromTemplate.Enabled && chkCreateBibleStudyNotebookFromTemplate.Checked)
                    || (chkCreateBibleCommentsNotebookFromTemplate.Enabled && chkCreateBibleCommentsNotebookFromTemplate.Checked)
                    || (chkCreateBibleNotesPagesNotebookFromTemplate.Enabled && chkCreateBibleNotesPagesNotebookFromTemplate.Checked)
                    || (chkCreateSingleNotebookFromTemplate.Enabled && chkCreateSingleNotebookFromTemplate.Checked))
                {
                    tbcMain.SelectedTab = tbcMain.TabPages[tpNotebooks.Name];
                }

                if (rbSingleNotebook.Checked && (module.UseSingleNotebook() || IsModerator || SettingsManager.Instance.IsSingleNotebook))
                {
                    SaveSingleNotebookParameters(module);
                }
                else
                {
                    SettingsManager.Instance.SectionGroupId_Bible = string.Empty;
                    SettingsManager.Instance.SectionGroupId_BibleStudy = string.Empty;
                    SettingsManager.Instance.SectionGroupId_BibleComments = string.Empty;
                    SettingsManager.Instance.SectionGroupId_BibleNotesPages = string.Empty;

                    if (_moduleWasChanged)
                    {
                        TryToSearchNotebooksForNewModule(module);                        
                        _moduleWasChanged = false;
                    }

                    SaveMultiNotebookParameters(module, ContainerType.Bible,
                        chkCreateBibleNotebookFromTemplate, cbBibleNotebook, BibleNotebookFromTemplatePath);

                    SaveMultiNotebookParameters(module, ContainerType.BibleStudy,
                        chkCreateBibleStudyNotebookFromTemplate, cbBibleStudyNotebook, BibleStudyNotebookFromTemplatePath);

                    SaveMultiNotebookParameters(module, ContainerType.BibleComments,
                        chkCreateBibleCommentsNotebookFromTemplate, cbBibleCommentsNotebook, BibleCommentsNotebookFromTemplatePath);

                    if (!chkUseFolderForBibleNotesPages.Checked)
                    {
                        //if (!CommitChangesAfterLoad)
                        //    ShownMessagesManager.SetMessageWasShown(ShownMessagesManager.MessagesCodes.SuggestUsingFolderForNotesPages);

                        SaveMultiNotebookParameters(module, ContainerType.BibleNotesPages,
                            chkCreateBibleNotesPagesNotebookFromTemplate, cbBibleNotesPagesNotebook, BibleNotesPagesNotebookFromTemplatePath);                        
                    }
                    else
                    {
                        if (string.IsNullOrEmpty(tbBibleNotesPagesFolder.Text))
                            SettingsManager.Instance.FolderPath_BibleNotesPages = Utils.GetNotesPagesFolderPath();  // почему-то иногда в текстбоксе оказывается пустое значение.
                        else
                            SettingsManager.Instance.FolderPath_BibleNotesPages = tbBibleNotesPagesFolder.Text;

                        SettingsManager.Instance.NotebookId_BibleNotesPages = string.Empty;
                        NotesPageManagerFS.UpdateResources();                        
                    }
                }

                this.TopMost = false;  // нам не нужен уже топ мост, потому что раньше он нам нужен был из-за того, что OneNote постоянно перекрывал программу когда создавались новые записные книжки

                if (!FormLogger.WasErrorLogged)
                {
                    SetProgramParameters();
                    SettingsManager.Instance.Save();
                    
                    RefreshCache(); // сразу обновляем кэш                    
                }

                if (!string.IsNullOrEmpty(SettingsManager.Instance.NotebookId_Bible))
                {
                    if (!BibleVersesLinksCacheManager.CacheIsActive(SettingsManager.Instance.NotebookId_Bible) && !ToIndexBible && !NotAskToIndexBible)
                    {
                        var minutes = GetMinutesForBibleVersesCacheGenerating();
                        if (MessageBox.Show(string.Format(BibleCommon.Resources.Constants.IndexBibleQuestion, minutes), BibleCommon.Resources.Constants.Warning,
                                        MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                            ToIndexBible = true;
                    }
                }
                
                if (ToIndexBible)
                {
                    IndexBible();
                    ToIndexBible = false;
                    _needToRefreshCache = true;
                }

                if (TryToIndexUnindexedDictionaries())
                    _needToRefreshCache = true;

                if (_needToRefreshCache)
                {
                    RefreshCache();
                    _needToRefreshCache = false;
                }

                if (!FormLogger.WasErrorLogged)
                {   
                    if (closeForm)
                        Close();
                    else
                    {
                        ReLoadParameters(false);
                        _originalModuleShortName = SettingsManager.Instance.ModuleShortName;
                    }
                }
            }
            catch (SaveParametersException ex)
            {
                FormLogger.LogError(ex);
                if (ex.NeedToReload)
                    LoadParameters(module, null);

                lblWarning.Visible = lblWarningVisibilityBefore;
                tbcMain.SelectedTab = tbcMain.TabPages[tpNotebooks.Name];
            }
            finally
            {
                btnOK.Enabled = true;                
                btnApply.Enabled = true;
                this.TopMost = false;
            }
        }

        internal static int GetMinutesForBibleVersesCacheGenerating()
        {
            return !SettingsManager.Instance.UseProxyLinksForBibleVerses || SettingsManager.Instance.GenerateFullBibleVersesCache ? 30 : 5;
        }

        private void RefreshCache()
        {
            if (_oneNoteApp.Windows.CurrentWindow != null)            
                Process.Start(RefreshCacheHandler.GetCommandUrlStatic(RefreshCacheHandler.RefreshCacheMode.RefreshApplicationCache));   // если текущее окно закрыто, то и кэш скорее всего закрыт. Когда окно откроют, кэш обновится                            
        }

        private void TryToSearchNotebooksForNewModule(ModuleInfo module)
        {
            var notebooks = OneNoteUtils.GetExistingNotebooks(ref _oneNoteApp);
            
            //notebooks.Remove(SettingsManager.Instance.NotebookId_Bible);
            //notebooks.Remove(SettingsManager.Instance.NotebookId_BibleStudy);
            //notebooks.Remove(SettingsManager.Instance.NotebookId_BibleComments);
            //if (!string.IsNullOrEmpty(SettingsManager.Instance.NotebookId_BibleNotesPages) && notebooks.ContainsKey(SettingsManager.Instance.NotebookId_BibleNotesPages))
            //    notebooks.Remove(SettingsManager.Instance.NotebookId_BibleNotesPages);

            ApplicationCache.Instance.RefreshHierarchyCache();

            TryToSearchNotebookForNewModule(module, ContainerType.Bible, SettingsManager.Instance.NotebookId_Bible,
                chkCreateBibleNotebookFromTemplate, cbBibleNotebook, ref notebooks, null);

            var commentsNotebookId = TryToSearchNotebookForNewModule(module, ContainerType.BibleComments, SettingsManager.Instance.NotebookId_BibleComments,
                chkCreateBibleCommentsNotebookFromTemplate, cbBibleCommentsNotebook, ref notebooks, null);

            if (!SettingsManager.Instance.StoreNotesPagesInFolder)
            {
                TryToSearchNotebookForNewModule(module, ContainerType.BibleNotesPages, SettingsManager.Instance.NotebookId_BibleNotesPages,
                    chkCreateBibleNotesPagesNotebookFromTemplate, cbBibleNotesPagesNotebook, ref notebooks, commentsNotebookId);
            }
        }

        private string TryToSearchNotebookForNewModule(ModuleInfo module, ContainerType containerType, string currentNotebookId,
            CheckBox chkCreateNotebookFromTemplate, ComboBox cbNotebook, ref Dictionary<string, string> notebooks, string defaultNotebookId)
        {
            if (!chkCreateNotebookFromTemplate.Checked)
            {
                string notebookName;
                var notebookId = GetNotebookIdFromCombobox(cbNotebook, out notebookName);
                if (notebookId == currentNotebookId)  // то есть если пользователь уже сам поменял, то не трогаем
                {                    
                    notebookId = SearchForNotebook(module, notebooks.Keys, containerType);

                    if (string.IsNullOrEmpty(notebookId) && !string.IsNullOrEmpty(defaultNotebookId))
                        notebookId = defaultNotebookId;

                    if (string.IsNullOrEmpty(notebookId))
                        chkCreateNotebookFromTemplate.Checked = true;
                    else
                    {
                        cbNotebook.SelectedItem = notebookId;
                        notebooks.Remove(notebookId);
                        return notebookId;
                    }
                }
            }

            return null;
        }

        private bool AreThereUnindexedDictionaries()
        {
            try
            {
                var modulesToDelete = new List<StoredModuleInfo>();
                foreach (var dictionaryInfo in SettingsManager.Instance.DictionariesModules)
                {
                    try
                    {
                        if (!DictionaryTermsCacheManager.CacheIsActive(dictionaryInfo.ModuleName))
                        {
                            var moduleInfo = ModulesManager.GetModuleInfo(dictionaryInfo.ModuleName);
                            if (moduleInfo != null && moduleInfo.Type != ModuleType.Strong)
                                return true;
                        }
                    }
                    catch (ModuleNotFoundException)
                    {
                        modulesToDelete.Add(dictionaryInfo);
                    }
                    catch (InvalidModuleException ex)
                    {
                        FormLogger.LogError(ex);
                        DeleteModuleWithConfirm(dictionaryInfo.ModuleName, false);
                        modulesToDelete.Add(dictionaryInfo);
                    }
                }

                foreach (var moduleInfo in modulesToDelete)
                {
                    SettingsManager.Instance.DictionariesModules.Remove(moduleInfo);
                }
            }
            catch (Exception ex)
            {
                FormLogger.LogError(ex);
            }

            return false;
        }

        private bool TryToIndexUnindexedDictionaries()
        {
            var result = false;

            foreach (var dictionaryInfo in SettingsManager.Instance.DictionariesModules.ToArray())
            {
                if (!DictionaryTermsCacheManager.CacheIsActive(dictionaryInfo.ModuleName) || dictionaryInfo.ModuleName == ForceIndexDictionaryModuleName)
                {
                    try
                    {
                        var moduleInfo = ModulesManager.GetModuleInfo(dictionaryInfo.ModuleName);
                        PrepareForLongProcessing(moduleInfo.NotebooksStructure.DictionaryTermsCount.Value, 1, BibleCommon.Resources.Constants.AddDictionaryStart);
                        LongProcessLogger.Preffix = string.Format("{0}: ", BibleCommon.Resources.Constants.IndexDictionary);
                        List<string> notFoundTerms;
                        DictionaryTermsCacheManager.GenerateCache(ref _oneNoteApp, moduleInfo, LongProcessLogger, out notFoundTerms);
                        LongProcessingDone(BibleCommon.Resources.Constants.AddDictionaryFinishMessage);
                        result = true;

                        if (notFoundTerms != null && notFoundTerms.Count > 0)
                        {
                            using (var form = new ErrorsForm())
                            {
                                form.AllErrors.Add(new ErrorsList(notFoundTerms)
                                {
                                    ErrorsDecription = BibleCommon.Resources.Constants.DictionaryTermsNotFound
                                });
                                form.ShowDialog();
                            }
                        }
                    }
                    catch (ProcessAbortedByUserException)
                    {
                        BibleCommon.Services.Logger.LogMessage("Process aborted by user");
                        LongProcessingDone(BibleCommon.Resources.Constants.ProcessAbortedByUser);
                    }
                    catch (COMException ex)
                    {
                        if (OneNoteUtils.IsError(ex, Error.hrObjectDoesNotExist))
                        {
                            SettingsManager.Instance.DictionariesModules.Remove(dictionaryInfo);
                            SettingsManager.Instance.Save();
                            LongProcessingDone(string.Empty);
                        }
                        else
                            throw;
                    }                    
                }
            }

            return result;
        }

        private void IndexBible()
        {
            int chaptersCount = ModulesManager.GetBibleChaptersCount(SettingsManager.Instance.ModuleShortName, true);
            PrepareForLongProcessing(chaptersCount, 1, BibleCommon.Resources.Constants.IndexBibleStart);
            LongProcessLogger.Preffix = string.Format("{0}: ", BibleCommon.Resources.Constants.IndexBible);
            BibleVersesLinksCacheManager.GenerateBibleVersesLinks(ref _oneNoteApp,
                SettingsManager.Instance.NotebookId_Bible, SettingsManager.Instance.SectionGroupId_Bible, 
                !SettingsManager.Instance.UseProxyLinksForBibleVerses || SettingsManager.Instance.GenerateFullBibleVersesCache,
                LongProcessLogger);
            LongProcessingDone(BibleCommon.Resources.Constants.IndexBibleFinish);
        }

        private void SaveMultiNotebookParameters(ModuleInfo module, ContainerType notebookType,
            CheckBox createFromTemplateControl, ComboBox selectedNotebookNameControl, string notebookFromTemplatePath)
        {
            if (createFromTemplateControl.Checked)
            {
                var notebookInfo = module.GetNotebook(notebookType);
                string notebookTemplateFileName = notebookInfo.Name;
                string notebookFolderPath;
                string notebookName = CreateNotebookFromTemplate(notebookTemplateFileName, notebookFromTemplatePath, out notebookFolderPath);
                if (!string.IsNullOrEmpty(notebookName))
                {
                    var notebookId = WaitAndSaveParameters(module, notebookType, notebookFolderPath, notebookName, notebookInfo.Nickname, notebookFolderPath, notebookType == ContainerType.Bible);                         // выйдем из метода только когда OneNote отработает
                    createFromTemplateControl.Checked = false;  // чтоб если ошибки будут потом, он заново не создавал                    
                    selectedNotebookNameControl.Items.Add(new ComboBoxItem(notebookId, notebookInfo.GetNicknameSafe()));
                    selectedNotebookNameControl.SelectedItem = notebookId;
                }
            }
            else
            {
                string notebookName;
                var notebookId = GetNotebookIdFromCombobox(selectedNotebookNameControl, out notebookName);
                TryToSaveNotebookParameters(notebookType, notebookId, notebookName, false);
            }
        }

        private void SaveSingleNotebookParameters(ModuleInfo module)
        {
            string notebookId;
            string notebookName;

            if (chkCreateSingleNotebookFromTemplate.Checked)
            {
                var notebookInfo = module.GetNotebook(ContainerType.Single);
                string notebookTemplateFileName = notebookInfo.Name;
                string notebookFolderPath;
                notebookName = CreateNotebookFromTemplate(notebookTemplateFileName, SingleNotebookFromTemplatePath, out notebookFolderPath);
                if (!string.IsNullOrEmpty(notebookName))
                {
                    WaitAndSaveParameters(module, ContainerType.Single, notebookFolderPath, notebookName, notebookInfo.Nickname, notebookFolderPath, false);
                    SearchForCorrespondenceSectionGroups(module, SettingsManager.Instance.NotebookId_Bible);
                }
            }
            else
            {
                
                notebookId = GetNotebookIdFromCombobox(cbSingleNotebook, out notebookName);
                if (TryToSaveNotebookParameters(ContainerType.Single, notebookId, notebookName, false))
                {
                    if (_notebookParametersForm != null && _notebookParametersForm.RenamedSectionGroups.Count > 0)
                        RenameSectionGroupsForm(notebookId, _notebookParametersForm.RenamedSectionGroups);

                    if (!_wasSearchedSectionGroupsInSingleNotebook)
                    {
                        try
                        {
                            SearchForCorrespondenceSectionGroups(module, notebookId);
                        }
                        catch (InvalidNotebookException)
                        {
                            FormLogger.LogError(BibleCommon.Resources.Constants.ConfiguratorWrongNotebookSelected, notebookName, ContainerType.Single);
                        }
                    }
                }

            }
        }
        
        private void SetProgramParameters()
        {
            bool localeWasChanged = false;
            if (SettingsManager.Instance.Language != (int)((ComboBoxItem)cbLanguage.SelectedItem).Key)
            {
                localeWasChanged = true;
                SettingsManager.Instance.Language = (int)((ComboBoxItem)cbLanguage.SelectedItem).Key;                
            }

            if (chkDefaultParameters.Checked)
            {
                SettingsManager.Instance.UseDefaultSettings = true;                
            }
            else
            {
                if (WasModified())
                    SettingsManager.Instance.UseDefaultSettings = false;

                SaveIntegerSettings();
                SaveBooleanSettings();                
                SaveLocalazibleSettings(localeWasChanged);
            }  
        }

        private void SaveBooleanSettings()
        {
            SettingsManager.Instance.ExpandMultiVersesLinking = chkExpandMultiVersesLinking.Checked;
            SettingsManager.Instance.ExcludedVersesLinking = chkExcludedVersesLinking.Checked;
            SettingsManager.Instance.UseDifferentPagesForEachVerse = chkUseDifferentPages.Checked;
            SettingsManager.Instance.RubbishPage_Use = chkUseRubbishPage.Checked;
            SettingsManager.Instance.RubbishPage_ExpandMultiVersesLinking = chkRubbishExpandMultiVersesLinking.Checked;
            SettingsManager.Instance.RubbishPage_ExcludedVersesLinking = chkRubbishExcludedVersesLinking.Checked;
            SettingsManager.Instance.UseProxyLinksForStrong = chkUseProxyLinksForStrong.Checked;
            SettingsManager.Instance.UseProxyLinksForLinks = chkUseProxyLinksForLinks.Checked;

            if (SettingsManager.Instance.UseProxyLinksForBibleVerses != chkUseProxyLinksForBibleVerses.Checked && !chkUseProxyLinksForBibleVerses.Checked)  // то есть мы перестали использовать прокси ссылки для стихов Библии
            {
                if (BibleVersesLinksCacheManager.CacheIsActive(SettingsManager.Instance.NotebookId_Bible)) // здесь нельзя использовать ApplicationCache.Instance.IsBibleVersesLinksCacheActive, так как тот кэширует
                {
                    if (!ApplicationCache.Instance.BibleVersesLinksCacheContainsHyperLinks())  // если уже содержит ссылки, то не надо обновлять кэш
                        ApplicationCache.Instance.CleanBibleVersesLinksCache(false);
                }
            }

            SettingsManager.Instance.UseProxyLinksForBibleVerses = chkUseProxyLinksForBibleVerses.Checked;
        }

        private void SaveIntegerSettings()
        {
            if (!string.IsNullOrEmpty(tbNotesPageWidth.Text))
            {
                int notesPageWidth;
                if (!int.TryParse(tbNotesPageWidth.Text, out notesPageWidth) || notesPageWidth < 200 || notesPageWidth > 1000)
                    throw new SaveParametersException(string.Format("{0} '{1}'", BibleCommon.Resources.Constants.ConfiguratorWrongParameterValue, lblNotesPageWidth.Text), false);

                SettingsManager.Instance.PageWidth_Notes = notesPageWidth;
            }

            if (!string.IsNullOrEmpty(tbRubbishNotesPageWidth.Text))
            {
                int rubbishNotesPageWidth;
                if (!int.TryParse(tbRubbishNotesPageWidth.Text, out rubbishNotesPageWidth) || rubbishNotesPageWidth < 200 || rubbishNotesPageWidth > 1000)
                    throw new SaveParametersException(string.Format("{0} '{1}'", BibleCommon.Resources.Constants.ConfiguratorWrongParameterValue, lblRubbishNotesPageWidth.Text), false);
                SettingsManager.Instance.PageWidth_RubbishNotes = rubbishNotesPageWidth;
            }
        }

        private void SaveLocalazibleSettings(bool localeWasChanged)
        {
            CultureInfo resourceCulture = new CultureInfo(SettingsManager.Instance.Language);

            if (!string.IsNullOrEmpty(tbBookOverviewName.Text))
            {
                if (SettingsManager.Instance.SectionName_DefaultBookOverview == tbBookOverviewName.Text
                    && SettingsManager.Instance.SectionName_DefaultBookOverview == BibleCommon.Resources.Constants.DefaultPageNameDefaultBookOverview
                    && localeWasChanged)
                    SettingsManager.Instance.SectionName_DefaultBookOverview = BibleCommon.Resources.Constants.ResourceManager
                        .GetString(BibleCommon.Consts.Constants.ResourceName_DefaultPageNameDefaultBookOverview, resourceCulture);
                else
                    SettingsManager.Instance.SectionName_DefaultBookOverview = tbBookOverviewName.Text;
            }            

            if (!string.IsNullOrEmpty(tbCommentsPageName.Text))
            {
                if (SettingsManager.Instance.PageName_DefaultComments == tbCommentsPageName.Text
                    && SettingsManager.Instance.PageName_DefaultComments == BibleCommon.Resources.Constants.DefaultPageNameDefaultComments
                    && localeWasChanged)
                    SettingsManager.Instance.PageName_DefaultComments = BibleCommon.Resources.Constants.ResourceManager
                        .GetString(BibleCommon.Consts.Constants.ResourceName_DefaultPageNameDefaultComments, resourceCulture);
                else
                    SettingsManager.Instance.PageName_DefaultComments = tbCommentsPageName.Text;
            }

            if (!string.IsNullOrEmpty(tbNotesPageName.Text))
            {
                if (SettingsManager.Instance.PageName_Notes == tbNotesPageName.Text
                    && SettingsManager.Instance.PageName_Notes == BibleCommon.Resources.Constants.DefaultPageName_Notes
                    && localeWasChanged)
                    SettingsManager.Instance.PageName_Notes = BibleCommon.Resources.Constants.ResourceManager
                        .GetString(BibleCommon.Consts.Constants.ResourceName_DefaultPageName_Notes, resourceCulture);
                else
                    SettingsManager.Instance.PageName_Notes = tbNotesPageName.Text;
            }

            if (!string.IsNullOrEmpty(tbRubbishNotesPageName.Text))
            {
                if (SettingsManager.Instance.PageName_RubbishNotes == tbRubbishNotesPageName.Text
                    && SettingsManager.Instance.PageName_RubbishNotes == BibleCommon.Resources.Constants.DefaultPageName_RubbishNotes
                    && localeWasChanged)
                    SettingsManager.Instance.PageName_RubbishNotes = BibleCommon.Resources.Constants.ResourceManager
                        .GetString(BibleCommon.Consts.Constants.ResourceName_DefaultPageName_RubbishNotes, resourceCulture);                   
                else
                    SettingsManager.Instance.PageName_RubbishNotes = tbRubbishNotesPageName.Text;
            }
        }

        private bool WasModified()
        {
            return SettingsManager.Instance.SectionName_DefaultBookOverview != tbBookOverviewName.Text
                || SettingsManager.Instance.PageName_Notes != tbNotesPageName.Text
                || SettingsManager.Instance.PageName_DefaultComments != tbCommentsPageName.Text
                || SettingsManager.Instance.ExpandMultiVersesLinking != chkExpandMultiVersesLinking.Checked
                || SettingsManager.Instance.ExcludedVersesLinking != chkExcludedVersesLinking.Checked
                || SettingsManager.Instance.UseDifferentPagesForEachVerse != chkUseDifferentPages.Checked
                || SettingsManager.Instance.RubbishPage_Use != chkUseRubbishPage.Checked
                || SettingsManager.Instance.PageName_RubbishNotes != tbRubbishNotesPageName.Text
                || SettingsManager.Instance.RubbishPage_ExpandMultiVersesLinking != chkRubbishExpandMultiVersesLinking.Checked
                || SettingsManager.Instance.RubbishPage_ExcludedVersesLinking != chkRubbishExcludedVersesLinking.Checked
                || SettingsManager.Instance.PageWidth_Notes.ToString() != tbNotesPageWidth.Text
                || SettingsManager.Instance.PageWidth_RubbishNotes.ToString() != tbRubbishNotesPageWidth.Text
                || SettingsManager.Instance.UseProxyLinksForStrong != chkUseProxyLinksForStrong.Checked
                || SettingsManager.Instance.UseProxyLinksForLinks != chkUseProxyLinksForLinks.Checked
                || SettingsManager.Instance.UseProxyLinksForBibleVerses != chkUseProxyLinksForBibleVerses.Checked;

        }

        private string WaitAndSaveParameters(ModuleInfo module, ContainerType notebookType, string notebookPath, string notebookName, string notebookNickname, string notebookFolderPath, bool saveModuleInformationIntoFirstPage)
        {
            PrepareForLongProcessing(100, 1, string.Format("{0} '{1}'", BibleCommon.Resources.Constants.ConfiguratorNotebookCreation, notebookName));
            
            bool parametersWasLoad = false;
            string notebookId = null;                

            try
            {                
                for (int attemptNumber = 0; attemptNumber <= LoadParametersAttemptsCount; attemptNumber++)
                {
                    pbMain.PerformStep();
                    System.Windows.Forms.Application.DoEvents();

                    if (TryToSaveNotebookParameters(notebookType, notebookPath, true, out notebookId))
                    {
                        parametersWasLoad = true;
                        break;
                    }
                    else
                    {
                        if (attemptNumber > 5 && string.IsNullOrEmpty(notebookId))  // то есть прошло уже 25 секунд, а записная книжка даже ещё не создалась!!!
                        {
                            OneNoteUtils.UseOneNoteAPI(ref _oneNoteApp, () =>
                            {
                                _oneNoteApp.OpenHierarchy(notebookFolderPath, null, out notebookId);
                            });
                        }
                    }

                    var freq = 10;
                    for (var i = 0; i < freq; i++)
                    {
                        Thread.Sleep(LoadParametersPauseBetweenAttempts / freq);
                        System.Windows.Forms.Application.DoEvents();
                    }
                }

                if (!parametersWasLoad)
                    throw new SaveParametersException(BibleCommon.Resources.Constants.ConfiguratorCanNotRequestDataFromOneNote, true);
                else
                {
                    if (saveModuleInformationIntoFirstPage)
                    {
                        if (!string.IsNullOrEmpty(notebookId))
                            SaveModuleInformationIntoFirstPage(notebookId, module);
                    }

                    if (!string.IsNullOrEmpty(notebookNickname))
                        NotebookGenerator.TryToRenameNotebookSafe(ref _oneNoteApp, notebookId, notebookNickname);
                }

            }
            finally
            {
                LongProcessingDone(string.Empty);                
            }        

            return notebookId;
        }

       

        private void SaveModuleInformationIntoFirstPage(string notebookId, ModuleInfo module)
        {
            XmlNamespaceManager xnm;
            var firstNotebookPageEl = NotebookChecker.GetFirstNotebookBiblePageId(ref _oneNoteApp, notebookId, null, out xnm);
            if (firstNotebookPageEl != null)
            {
                var moduleInfo = new EmbeddedModuleInfo(module.ShortName, module.Version);
                var pageContent = OneNoteUtils.GetPageContent(ref _oneNoteApp, (string)firstNotebookPageEl.Attribute("ID"), out xnm);
                OneNoteUtils.UpdateElementMetaData(pageContent.Root, BibleCommon.Consts.Constants.Key_EmbeddedBibleModule, moduleInfo.ToString(), xnm);
                OneNoteUtils.UpdatePageContentSafe(ref _oneNoteApp, pageContent, xnm);
            }
        }

        private bool TryToSaveNotebookParameters(ContainerType notebookType, string notebookFolderPath, bool silientMode, out string notebookId)
        {                
            string notebookName;
            notebookId = OneNoteUtils.GetNotebookIdByPath(ref _oneNoteApp, notebookFolderPath, true, out notebookName);

            if (!string.IsNullOrEmpty(notebookId))
                return TryToSaveNotebookParameters(notebookType, notebookId, notebookName, silientMode);

            return false;
        }

        private bool TryToSaveNotebookParameters(ContainerType notebookType, string notebookId, string notebookName, bool silientMode)
        {
            try
            {
                var module = ModulesManager.GetCurrentModuleInfo();

                string errorText;
                if (NotebookChecker.CheckNotebook(ref _oneNoteApp, module, notebookId, notebookType, true, out errorText))
                {
                    switch (notebookType)
                    {
                        case ContainerType.Single:
                            SettingsManager.Instance.NotebookId_Bible = notebookId;
                            SettingsManager.Instance.NotebookId_BibleComments = notebookId;
                            SettingsManager.Instance.NotebookId_BibleNotesPages = notebookId;
                            SettingsManager.Instance.NotebookId_BibleStudy = notebookId;
                            break;
                        case ContainerType.Bible:
                            SettingsManager.Instance.NotebookId_Bible = notebookId;
                            break;
                        case ContainerType.BibleComments:
                        case ContainerType.BibleStudy:
                            {
                                if (notebookType == ContainerType.BibleComments)
                                    SettingsManager.Instance.NotebookId_BibleComments = notebookId;
                                else
                                    SettingsManager.Instance.NotebookId_BibleStudy = notebookId;

                                var notebookIdLocal = notebookId;
                                if (SettingsManager.Instance.SelectedNotebooksForAnalyze != null 
                                    && !SettingsManager.Instance.SelectedNotebooksForAnalyze.Exists(notebook => notebook.NotebookId == notebookIdLocal))
                                    SettingsManager.Instance.SelectedNotebooksForAnalyze.Add(new NotebookForAnalyzeInfo(notebookId));
                            }
                            break;
                        case ContainerType.BibleNotesPages:
                            SettingsManager.Instance.NotebookId_BibleNotesPages = notebookId;
                            break;
                    }

                    return true;
                }
                else
                {

                    string message = !string.IsNullOrEmpty(notebookId)
                                            ? string.Format(BibleCommon.Resources.Constants.ConfiguratorWrongNotebookSelected + "\n" + errorText, notebookName,
                                                                        ContainerTypeHelper.GetContainerTypeName(notebookType))
                                            : string.Format(BibleCommon.Resources.Constants.ConfiguratorNotebookNotDefinedForType, ContainerTypeHelper.GetContainerTypeName(notebookType));                                
                    
                    if (!silientMode)
                        throw new SaveParametersException(message, false);  
                    else
                        BibleCommon.Services.Logger.LogError(message);
                }
            }
            catch (Exception ex)
            {
                if (!silientMode)
                    throw new SaveParametersException(OneNoteUtils.ParseErrorAndMakeItMoreUserFriendly(ex.Message), false);
                else
                    BibleCommon.Services.Logger.LogError(ex);
            }

            return false;
        }

        private void SearchForCorrespondenceSectionGroups(ModuleInfo module, string notebookId)
        {
            ApplicationCache.HierarchyElement notebook = ApplicationCache.Instance.GetHierarchy(ref _oneNoteApp, notebookId, HierarchyScope.hsSections, true);

            List<ContainerType> sectionGroups = new List<ContainerType>();

            foreach (XElement sectionGroup in notebook.Content.Root.XPathSelectElements("one:SectionGroup", notebook.Xnm).Where(sg => !OneNoteUtils.IsRecycleBin(sg)))
            {
                string id = (string)sectionGroup.Attribute("ID");

                if (NotebookChecker.ElementIsBible(module, sectionGroup, notebook.Xnm) && !sectionGroups.Contains(ContainerType.Bible))
                {
                    SettingsManager.Instance.SectionGroupId_Bible = id;
                    sectionGroups.Add(ContainerType.Bible);
                }
                else if (NotebookChecker.ElementIsBibleComments(module, sectionGroup, notebook.Xnm) && !sectionGroups.Contains(ContainerType.BibleComments))
                {
                    SettingsManager.Instance.SectionGroupId_BibleComments = id;
                    SettingsManager.Instance.SectionGroupId_BibleNotesPages = id;
                    sectionGroups.Add(ContainerType.BibleComments);
                }
                else if (!sectionGroups.Contains(ContainerType.BibleStudy))
                {
                    SettingsManager.Instance.SectionGroupId_BibleStudy = id;
                    sectionGroups.Add(ContainerType.BibleStudy);
                }              
                else
                    throw new InvalidNotebookException();
            }

            if (sectionGroups.Count < 3)
                throw new InvalidNotebookException();
        }

        private void RenameSectionGroupsForm(string notebookId, Dictionary<string, string> renamedSectionGroups)
        {
            ApplicationCache.HierarchyElement notebook = ApplicationCache.Instance.GetHierarchy(ref _oneNoteApp, notebookId, HierarchyScope.hsSections, true);     

            foreach (string sectionGroupId in renamedSectionGroups.Keys)
            {
                XElement sectionGroup = notebook.Content.Root.XPathSelectElement(string.Format("one:SectionGroup[@ID=\"{0}\"]", sectionGroupId), notebook.Xnm);

                if (sectionGroup != null)
                {
                    sectionGroup.SetAttributeValue("name", renamedSectionGroups[sectionGroupId]);
                }
                else
                    FormLogger.LogError(string.Format("{0} '{1}'.", BibleCommon.Resources.Constants.ConfiguratorSectionGroupNotFound, sectionGroupId));
            }

            OneNoteUtils.UseOneNoteAPI(ref _oneNoteApp, () =>
            {
                _oneNoteApp.UpdateHierarchy(notebook.Content.ToString(), Constants.CurrentOneNoteSchema);
            });
            ApplicationCache.Instance.RefreshHierarchyCache(ref _oneNoteApp, notebookId, HierarchyScope.hsSections);     
        }

        private string CreateNotebookFromTemplate(string notebookTemplateFileName, string notebookFromTemplatePath, out string notebookFolderPath)
        {
            string s = null;
            notebookFolderPath = null;
            string packageDirectory = ModulesManager.GetCurrentModuleDirectiory();                
            string packageFilePath = Path.Combine(packageDirectory, notebookTemplateFileName);

            if (File.Exists(packageFilePath))
            {
                notebookFolderPath = Path.Combine(notebookFromTemplatePath, Path.GetFileNameWithoutExtension(notebookTemplateFileName));

                notebookFolderPath = Utils.GetNewDirectoryPath(notebookFolderPath);

                //if (!string.IsNullOrEmpty(folderPath))
                //{
                string notebookFolderPathTemp = notebookFolderPath;
                OneNoteUtils.UseOneNoteAPI(ref _oneNoteApp, () =>
                {
                    _oneNoteApp.OpenPackage(packageFilePath, notebookFolderPathTemp, out s);
                });

                string[] files = Directory.GetFiles(s, "*.onetoc2", SearchOption.TopDirectoryOnly);
                if (files.Length > 0)
                    Process.Start(files[0]);
                else
                    FormLogger.LogError(string.Format("{0} '{1}'.", BibleCommon.Resources.Constants.ConfiguratorErrorWhileNotebookOpenning, notebookTemplateFileName));

                return Path.GetFileNameWithoutExtension(notebookFolderPath);
                //}
                //else
                //    Logger.LogError(BibleCommon.Resources.Constants.ConfiguratorSelectAnotherFolder);
            }
            else
                FormLogger.LogError(string.Format("{0} '{1}'.", BibleCommon.Resources.Constants.ConfiguratorNotebookTemplateNotFound, packageFilePath));

            return string.Empty;
        }       

        private LoadForm _loadForm;
        private bool _firstShown = true;
        private void MainForm_Shown(object sender, EventArgs e)
        {
            if (_firstShown)
            {
                try
                {
                    bool? needSaveSettings = null;

                    if (ShowModulesTabAtStartUp)
                    {                        
                        tbcMain.SelectedTab = tbcMain.TabPages[tpModules.Name];
                        _wasLoadedModulesInfo = false;                        

                        if (NeedToSaveChangesAfterLoadingModuleAtStartUp)
                            needSaveSettings = true;
                    }
                    else if (string.IsNullOrEmpty(SettingsManager.Instance.ModuleShortName))
                    {
                        var modules = ModulesManager.GetModules(true);                        
                        if (modules.Where(m => m.Type == ModuleType.Bible).Count() == 1)                        
                            SettingsManager.Instance.ModuleShortName = modules.First(m => m.Type == ModuleType.Bible).ShortName;                        
                    }
                    
                    PrepareFolderBrowsers();
                    SetNotebooksDefaultPaths();

                    if (!SettingsManager.Instance.CurrentModuleIsCorrect())
                        tbcMain.SelectedTab = tbcMain.TabPages[tpModules.Name];                    
                    else
                    {
                        var module = ModulesManager.GetCurrentModuleInfo();
                        LoadParameters(module, needSaveSettings);
                        _originalModuleShortName = SettingsManager.Instance.ModuleShortName;
                    }

                    if (!IsModerator)
                    {
                        btnConverter.Visible = false;
                        btnModuleChecker.Visible = false;
                    }

                    //chkUseProxyLinks.Visible = false;

                    this.Text += string.Format(" v{0}", SettingsManager.Instance.CurrentVersion);
                    this.SetFocus();
                    _firstShown = false;
                }                
                finally
                {
                    _loadForm.Hide();
                }

                if (CommitChangesAfterLoad)
                    btnOK_Click(this, null);
            }
        }      

        private void MainForm_Load(object sender, EventArgs e)
        {
            _loadForm = new LoadForm();

            _loadForm.Show();
        }

        private void LoadParameters(ModuleInfo module, bool? forceNeedToSaveSettings)
        {
            if (!SettingsManager.Instance.IsConfigured(ref _oneNoteApp) || forceNeedToSaveSettings.GetValueOrDefault(false) || AreThereUnindexedDictionaries())
                lblWarning.Visible = true;
            //else  // пусть лучше будет так, чтобы если пользователь что-то поменял - его программа просила всегда сохранить изменения, пока он не сохранит
            //    lblWarning.Visible = false;
            
            ApplicationCache.Instance.RefreshHierarchyCache();

            _notebooks = OneNoteUtils.GetExistingNotebooks(ref _oneNoteApp);
            string singleNotebookId = (IsModerator || module.UseSingleNotebook() || SettingsManager.Instance.IsSingleNotebook) ? SearchForNotebook(module, _notebooks.Keys, ContainerType.Single) : string.Empty;
            string bibleNotebookId = SearchForNotebook(module, _notebooks.Keys, ContainerType.Bible);
            string bibleCommentsNotebookId = SearchForNotebook(module, _notebooks.Keys, ContainerType.BibleComments);
            string bibleStudyNotebookId = SearchForNotebook(module, _notebooks.Keys, ContainerType.BibleStudy);
            string bibleNotesPagesNotebookId = SearchForNotebook(module, _notebooks.Keys.ToList().Where(s => s != bibleCommentsNotebookId), ContainerType.BibleNotesPages);
            if (string.IsNullOrEmpty(bibleNotesPagesNotebookId))
                bibleNotesPagesNotebookId = bibleCommentsNotebookId;

            if (string.IsNullOrEmpty(bibleNotebookId) && string.IsNullOrEmpty(bibleCommentsNotebookId))  // а то иначе он всегда "Личную" при установке выбирает
                bibleStudyNotebookId = null;

            rbSingleNotebook.Checked = SettingsManager.Instance.IsSingleNotebook 
                                    && !string.IsNullOrEmpty(singleNotebookId);

            rbMultiNotebook.Checked = !rbSingleNotebook.Checked;
            rbMultiNotebook_CheckedChanged(this, null);

            cbSingleNotebook.Items.Clear();
            cbBibleNotebook.Items.Clear();
            cbBibleCommentsNotebook.Items.Clear();
            cbBibleNotesPagesNotebook.Items.Clear();
            cbBibleStudyNotebook.Items.Clear();

            foreach (var notebookId in _notebooks.Keys)
            {
                var cbItem = new ComboBoxItem(notebookId, _notebooks[notebookId]);
                cbSingleNotebook.Items.Add(cbItem);
                cbBibleNotebook.Items.Add(cbItem);
                cbBibleCommentsNotebook.Items.Add(cbItem);
                cbBibleNotesPagesNotebook.Items.Add(cbItem);
                cbBibleStudyNotebook.Items.Add(cbItem);
            }

            var canUseSingleNotebook = module.UseSingleNotebook() || IsModerator || SettingsManager.Instance.IsSingleNotebook;

            if (canUseSingleNotebook)
            {
                var defaultNotebookName = "";
                if (module.UseSingleNotebook())
                    defaultNotebookName = Path.GetFileNameWithoutExtension(module.GetNotebook(ContainerType.Single).Name);
                SetNotebookParameters(rbSingleNotebook.Checked,
                    !string.IsNullOrEmpty(singleNotebookId) ? _notebooks[singleNotebookId] : defaultNotebookName,
                    _notebooks, SettingsManager.Instance.NotebookId_Bible, cbSingleNotebook, chkCreateSingleNotebookFromTemplate);
            }

            SetNotebookParameters(rbMultiNotebook.Checked, !string.IsNullOrEmpty(bibleNotebookId) ? _notebooks[bibleNotebookId] :
                Path.GetFileNameWithoutExtension(module.GetNotebook(ContainerType.Bible).Name),
                _notebooks, SettingsManager.Instance.NotebookId_Bible, cbBibleNotebook, chkCreateBibleNotebookFromTemplate);

            SetNotebookParameters(rbMultiNotebook.Checked, !string.IsNullOrEmpty(bibleStudyNotebookId) ? _notebooks[bibleStudyNotebookId] :
                Path.GetFileNameWithoutExtension(module.GetNotebook(ContainerType.BibleStudy).Name),
                _notebooks, SettingsManager.Instance.NotebookId_BibleStudy, cbBibleStudyNotebook, chkCreateBibleStudyNotebookFromTemplate);

            SetNotebookParameters(rbMultiNotebook.Checked, !string.IsNullOrEmpty(bibleCommentsNotebookId) ? _notebooks[bibleCommentsNotebookId] :
                Path.GetFileNameWithoutExtension(module.GetNotebook(ContainerType.BibleComments).Name),
                _notebooks, SettingsManager.Instance.NotebookId_BibleComments, cbBibleCommentsNotebook, chkCreateBibleCommentsNotebookFromTemplate);

            SetNotebookParameters(rbMultiNotebook.Checked,
                                      !string.IsNullOrEmpty(bibleNotesPagesNotebookId)
                                                    ? _notebooks[bibleNotesPagesNotebookId]
                                                    : Path.GetFileNameWithoutExtension(module.GetNotebook(ContainerType.BibleNotesPages).Name),
                                      _notebooks, SettingsManager.Instance.NotebookId_BibleNotesPages, cbBibleNotesPagesNotebook, chkCreateBibleNotesPagesNotebookFromTemplate);


            if (SettingsManager.Instance.StoreNotesPagesInFolder || !_notebooks.ContainsKey(SettingsManager.Instance.NotebookId_BibleNotesPages))
            {
                chkUseFolderForBibleNotesPages.Checked = true;
            }
            else if (rbMultiNotebook.Checked)
            {
                    chkUseFolderForBibleNotesPages.Checked = false;
            }

            tbBibleNotesPagesFolder.Text = SettingsManager.Instance.FolderPath_BibleNotesPages;
            ttNotesPageFolder.SetToolTip(tbBibleNotesPagesFolder, tbBibleNotesPagesFolder.Text);

            tbBookOverviewName.Text = SettingsManager.Instance.SectionName_DefaultBookOverview;
            tbNotesPageName.Text = SettingsManager.Instance.PageName_Notes;
            tbCommentsPageName.Text = SettingsManager.Instance.PageName_DefaultComments;
            tbNotesPageWidth.Text = SettingsManager.Instance.PageWidth_Notes.ToString();
            chkExpandMultiVersesLinking.Checked = SettingsManager.Instance.ExpandMultiVersesLinking;
            chkExcludedVersesLinking.Checked = SettingsManager.Instance.ExcludedVersesLinking;
            chkUseDifferentPages.Checked = SettingsManager.Instance.UseDifferentPagesForEachVerse;

            chkUseRubbishPage.Checked = SettingsManager.Instance.RubbishPage_Use;
            tbRubbishNotesPageName.Text = SettingsManager.Instance.PageName_RubbishNotes;
            tbRubbishNotesPageWidth.Text = SettingsManager.Instance.PageWidth_RubbishNotes.ToString();
            chkRubbishExpandMultiVersesLinking.Checked = SettingsManager.Instance.RubbishPage_ExpandMultiVersesLinking;
            chkRubbishExcludedVersesLinking.Checked = SettingsManager.Instance.RubbishPage_ExcludedVersesLinking;

            chkUseProxyLinksForStrong.Checked = SettingsManager.Instance.UseProxyLinksForStrong;
            chkUseProxyLinksForLinks.Checked = SettingsManager.Instance.UseProxyLinksForLinks;
            chkUseProxyLinksForBibleVerses.Checked = SettingsManager.Instance.UseProxyLinksForBibleVerses;

            chkUseRubbishPage_CheckedChanged(this, null);
            chkUseFolderForBibleNotesPages_CheckedChanged(this, null);

            InitLanguagesMenu();

            if (!rbSingleNotebook.Checked && !IsModerator)
                rbSingleNotebook.Enabled = false;            
        }        

        private void InitLanguagesMenu()
        {
            var languages = LanguageManager.GetDisplayedNames();

            var currentLanguage = LanguageManager.GetCurrentCultureInfo();

            cbLanguage.Items.Clear();
            foreach (var pair in languages)
            {
                cbLanguage.Items.Add(new ComboBoxItem(pair.Key, pair.Value));
                if (pair.Key == currentLanguage.LCID)
                    cbLanguage.SelectedIndex = cbLanguage.Items.Count - 1;

            }
        }

        private string SearchForNotebook(ModuleInfo module, IEnumerable<string> notebooksIds, ContainerType notebookType)
        {
            string errorText;
            if (notebookType == ContainerType.BibleStudy)
            {
                try
                {
                    var nickname = SettingsManager.Instance.CurrentModuleCached.NotebooksStructure.Notebooks.FirstOrDefault(n => n.Type == ContainerType.BibleStudy).Nickname;
                    foreach (var notebookId in notebooksIds)
                    {
                        var name = OneNoteUtils.GetNotebookElementNickname(ref _oneNoteApp, notebookId);
                        if (name == nickname && NotebookChecker.CheckNotebook(ref _oneNoteApp, module, notebookId, notebookType, false, out errorText))
                            return notebookId;
                    }
                }
                catch (Exception ex)                // наверное зря, но просто добавляю этот код перед самым релизом, потому опасаюсь.
                {
                    FormLogger.LogError(ex);
                }
            }

            foreach (var notebookId in notebooksIds)
            {                
                if (NotebookChecker.CheckNotebook(ref _oneNoteApp, module, notebookId, notebookType, false, out errorText))                
                    return notebookId;                
            }

            return null;
        }

        private static void SetNotebookParameters(bool loadNameFromSettings, string defaultName, Dictionary<string, string> notebooks, 
            string notebookIdFromSettings, ComboBox cb, CheckBox chk)
        {
            chk.Checked = false;
            var notebookId = loadNameFromSettings ? notebookIdFromSettings : string.Empty;
            if (!string.IsNullOrEmpty(notebookId) && cb.Items.Contains(notebookId))
                cb.SelectedItem = notebookId;
            else
            {
                var defaultNotebook = cb.Items.Cast<ComboBoxItem>().FirstOrDefault(item => item.Value.IndexOf(defaultName) > -1);
                if (defaultNotebook != null)
                    cb.SelectedItem = defaultNotebook;
                else
                    chk.Checked = true;
            }
        }      

        private void SetNotebooksDefaultPaths()
        {
            // по дефолту пути такие
            SingleNotebookFromTemplatePath = folderBrowserDialog.SelectedPath;
            BibleNotebookFromTemplatePath = folderBrowserDialog.SelectedPath;
            BibleCommentsNotebookFromTemplatePath = folderBrowserDialog.SelectedPath;
            BibleNotesPagesNotebookFromTemplatePath = folderBrowserDialog.SelectedPath;
            BibleStudyNotebookFromTemplatePath = folderBrowserDialog.SelectedPath;
        }

        private void PrepareFolderBrowsers()
        {
            string defaultNotebookFolderPath = null;

            OneNoteUtils.UseOneNoteAPI(ref _oneNoteApp, () =>
            {
                _oneNoteApp.GetSpecialLocation(SpecialLocation.slDefaultNotebookFolder, out defaultNotebookFolderPath);
            });

            
            folderBrowserDialog.SelectedPath = defaultNotebookFolderPath;
            folderBrowserDialog.Description = BibleCommon.Resources.Constants.ConfiguratorSetNotebookFolder;

            string toolTipMessage = BibleCommon.Resources.Constants.DefineNotebookDirectory;
            FormExtensions.SetToolTip(btnSingleNotebookSetPath, toolTipMessage);
            FormExtensions.SetToolTip(btnBibleNotebookSetPath, toolTipMessage);
            FormExtensions.SetToolTip(btnBibleStudyNotebookSetPath, toolTipMessage);
            FormExtensions.SetToolTip(btnBibleCommentsNotebookSetPath, toolTipMessage);
            FormExtensions.SetToolTip(btnBibleNotesPagesNotebookSetPath, toolTipMessage);

            notesPagesFolderBrowserDialog.SelectedPath = SettingsManager.Instance.FolderPath_BibleNotesPages;
            notesPagesFolderBrowserDialog.Description = BibleCommon.Resources.Constants.ConfiguratorSetFolderForNotesPages;
            FormExtensions.SetToolTip(btnBibleNotesPagesSetFolder, BibleCommon.Resources.Constants.ConfiguratorSetFolderForNotesPages);            
        }

      

        private void rbMultiNotebook_CheckedChanged(object sender, EventArgs e)
        {
            cbSingleNotebook.Enabled = rbSingleNotebook.Checked;
            lblSelectSingleNotebook.Enabled = rbSingleNotebook.Checked;
            btnSingleNotebookParameters.Enabled = rbSingleNotebook.Checked;
            chkCreateSingleNotebookFromTemplate.Enabled = false; // rbSingleNotebook.Checked;
            btnSingleNotebookParameters.Enabled = rbSingleNotebook.Checked;
            btnSingleNotebookSetPath.Enabled = rbSingleNotebook.Checked;

            cbBibleNotebook.Enabled = rbMultiNotebook.Checked;
            cbBibleCommentsNotebook.Enabled = rbMultiNotebook.Checked;
            cbBibleNotesPagesNotebook.Enabled = rbMultiNotebook.Checked;
            cbBibleStudyNotebook.Enabled = rbMultiNotebook.Checked;
            chkCreateBibleNotebookFromTemplate.Enabled = rbMultiNotebook.Checked;
            chkCreateBibleCommentsNotebookFromTemplate.Enabled = rbMultiNotebook.Checked;
            chkCreateBibleNotesPagesNotebookFromTemplate.Enabled = rbMultiNotebook.Checked;
            chkCreateBibleStudyNotebookFromTemplate.Enabled = rbMultiNotebook.Checked;
            btnBibleNotebookSetPath.Enabled = rbMultiNotebook.Checked;
            btnBibleCommentsNotebookSetPath.Enabled = rbMultiNotebook.Checked;
            btnBibleNotesPagesNotebookSetPath.Enabled = rbMultiNotebook.Checked;
            btnBibleStudyNotebookSetPath.Enabled = rbMultiNotebook.Checked;

            chkUseFolderForBibleNotesPages.Enabled = rbMultiNotebook.Checked;
            tbBibleNotesPagesFolder.Enabled = rbMultiNotebook.Checked;
            btnBibleNotesPagesSetFolder.Enabled = rbMultiNotebook.Checked;

            if (rbSingleNotebook.Checked)
            {
                chkCreateSingleNotebookFromTemplate_CheckedChanged(this, null);
            }
            else
            {
                chkCreateBibleNotebookFromTemplate_CheckedChanged(this, null);
                chkCreateBibleCommentsNotebookFromTemplate_CheckedChanged(this, null);
                chkCreateBibleStudyNotebookFromTemplate_CheckedChanged(this, null);

                SetNotesPagesNotebookControlsAbility();                                
            }

            SetNotesPageAnalyzeControlsAbility();
            SetRubbishControlsAbility();                
        }

        private void chkCreateSingleNotebookFromTemplate_CheckedChanged(object sender, EventArgs e)
        {
            cbSingleNotebook.Enabled = !chkCreateSingleNotebookFromTemplate.Checked; // && chkCreateSingleNotebookFromTemplate.Enabled;
            btnSingleNotebookParameters.Enabled = !chkCreateSingleNotebookFromTemplate.Checked; // && chkCreateSingleNotebookFromTemplate.Enabled;
            btnSingleNotebookSetPath.Enabled = chkCreateSingleNotebookFromTemplate.Checked; // && chkCreateSingleNotebookFromTemplate.Enabled;
        }

        private void chkCreateBibleNotebookFromTemplate_CheckedChanged(object sender, EventArgs e)
        {
            cbBibleNotebook.Enabled = chkCreateBibleNotebookFromTemplate.Enabled && !chkCreateBibleNotebookFromTemplate.Checked;
            btnBibleNotebookSetPath.Enabled = chkCreateBibleNotebookFromTemplate.Enabled && chkCreateBibleNotebookFromTemplate.Checked;
        }

        private void chkCreateBibleCommentsNotebookFromTemplate_CheckedChanged(object sender, EventArgs e)
        {
            cbBibleCommentsNotebook.Enabled = chkCreateBibleCommentsNotebookFromTemplate.Enabled && !chkCreateBibleCommentsNotebookFromTemplate.Checked;
            btnBibleCommentsNotebookSetPath.Enabled = chkCreateBibleCommentsNotebookFromTemplate.Enabled && chkCreateBibleCommentsNotebookFromTemplate.Checked;
        }

        private void chkCreateBibleNotesPagesNotebookFromTemplate_CheckedChanged(object sender, EventArgs e)
        {
            SetNotesPagesNotebookControlsAbility();
        }

        private void chkCreateBibleStudyNotebookFromTemplate_CheckedChanged(object sender, EventArgs e)
        {
            cbBibleStudyNotebook.Enabled = chkCreateBibleStudyNotebookFromTemplate.Enabled && !chkCreateBibleStudyNotebookFromTemplate.Checked;
            btnBibleStudyNotebookSetPath.Enabled = chkCreateBibleStudyNotebookFromTemplate.Enabled && chkCreateBibleStudyNotebookFromTemplate.Checked;
        }

        private static string GetNotebookIdFromCombobox(ComboBox cb, out string notebookName)
        {
            notebookName = null;
            if (cb.SelectedItem != null)
            {
                notebookName = cb.SelectedItem.ToString();
                return ((ComboBoxItem)cb.SelectedItem).Key.ToString();
            }

            return null;
        }

        private void btnSingleNotebookParameters_Click(object sender, EventArgs e)
        {   
            try
            {
                string notebookName;
                var notebookId = GetNotebookIdFromCombobox(cbSingleNotebook, out notebookName);
                if (!string.IsNullOrEmpty(notebookId))
                {                   
                    
                var module = ModulesManager.GetCurrentModuleInfo();
                string errorText;
                    if (NotebookChecker.CheckNotebook(ref _oneNoteApp, module, notebookId, ContainerType.Single, true, out errorText))
                {
                    if (_notebookParametersForm == null)
                        _notebookParametersForm = new NotebookParametersForm(notebookId);

                    if (_notebookParametersForm.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {   
                        SettingsManager.Instance.SectionGroupId_Bible = _notebookParametersForm.GroupedSectionGroups[ContainerType.Bible];
                        SettingsManager.Instance.SectionGroupId_BibleStudy = _notebookParametersForm.GroupedSectionGroups[ContainerType.BibleStudy];
                        SettingsManager.Instance.SectionGroupId_BibleComments = _notebookParametersForm.GroupedSectionGroups[ContainerType.BibleComments];
                        SettingsManager.Instance.SectionGroupId_BibleNotesPages = _notebookParametersForm.GroupedSectionGroups[ContainerType.BibleComments];

                            _wasSearchedSectionGroupsInSingleNotebook = true;  // нашли необходимые группы разделов. 
                    }
                }
                else
                {
                        FormLogger.LogError(string.Format(BibleCommon.Resources.Constants.ConfiguratorWrongNotebookSelected + "\n" + errorText, notebookName, 
                            ContainerTypeHelper.GetContainerTypeName(ContainerType.Single)));
                }
            }
            else
            {
                FormLogger.LogMessage(BibleCommon.Resources.Constants.ConfiguratorNotebookNotDefined);
            }
        }
            catch (Exception ex)
            {
                FormLogger.LogError(ex);
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
            }
        }

        private void btnBibleNotesPagesNotebookSetPath_Click(object sender, EventArgs e)
        {
            if (chkCreateBibleNotesPagesNotebookFromTemplate.Checked)
            {
                if (folderBrowserDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    BibleNotesPagesNotebookFromTemplatePath = folderBrowserDialog.SelectedPath;
                }
            }
        }

        private void chkDefaultPageNameParameters_CheckedChanged(object sender, EventArgs e)
        {
            tbCommentsPageName.Enabled = !chkDefaultParameters.Checked;
            tbNotesPageName.Enabled = !chkDefaultParameters.Checked;
            tbBookOverviewName.Enabled = !chkDefaultParameters.Checked;
            tbNotesPageWidth.Enabled = !chkDefaultParameters.Checked;
            chkExpandMultiVersesLinking.Enabled = !chkDefaultParameters.Checked;
            chkExcludedVersesLinking.Enabled = !chkDefaultParameters.Checked;
            chkUseDifferentPages.Enabled = !chkDefaultParameters.Checked;
            chkUseRubbishPage.Enabled = !chkDefaultParameters.Checked;
            tbRubbishNotesPageName.Enabled = !chkDefaultParameters.Checked;
            tbRubbishNotesPageWidth.Enabled = !chkDefaultParameters.Checked;
            chkRubbishExpandMultiVersesLinking.Enabled = !chkDefaultParameters.Checked;
            chkRubbishExcludedVersesLinking.Enabled = !chkDefaultParameters.Checked;
            chkUseProxyLinksForStrong.Enabled = !chkDefaultParameters.Checked;
            chkUseProxyLinksForBibleVerses.Enabled = !chkDefaultParameters.Checked;
            chkUseProxyLinksForLinks.Enabled = !chkDefaultParameters.Checked;

            chkUseRubbishPage_CheckedChanged(this, null);
            chkUseFolderForBibleNotesPages_CheckedChanged(this, null);
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            StopLongProcess = true;
            LongProcessLogger.AbortedByUser = true;            
        }

        private void MainForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            BibleCommon.Services.Logger.Done();
            OneNoteUtils.ReleaseOneNoteApp(ref _oneNoteApp);

            if (_notebookParametersForm != null)
                _notebookParametersForm.Dispose();

            if (_loadForm != null)
                _loadForm.Dispose();
        }

        private void btnRelinkComments_Click(object sender, EventArgs e)
        {
            using (var manager = new RelinkAllBibleCommentsManager(this))
            {
                manager.RelinkAllBibleComments();
            }
        }

        public void PrepareForLongProcessing(int pbMaxValue, int pbStep, string infoText)
        {
            pbMain.Value = 0;
            pbMain.Maximum = pbMaxValue;
            pbMain.Step = pbStep;
            pbMain.Visible = true;

            tbcMain.Enabled = false;
            lblProgressInfo.Text = infoText;

            btnOK.Enabled = false;            
            btnApply.Enabled = false;
            System.Windows.Forms.Application.DoEvents();
        }

        public void LongProcessingDone(string infoText)
        {
            FormExtensions.SetControlPropertyThreadSafe(pbMain, "Value", 0);
            FormExtensions.SetControlPropertyThreadSafe(pbMain, "Maximum", 100);
            FormExtensions.SetControlPropertyThreadSafe(pbMain, "Step", 1);
            FormExtensions.SetControlPropertyThreadSafe(pbMain, "Visible", false);
            FormExtensions.SetControlPropertyThreadSafe(tbcMain, "Enabled", true);
            FormExtensions.SetControlPropertyThreadSafe(lblProgressInfo, "Text", infoText);
            FormExtensions.SetControlPropertyThreadSafe(btnOK, "Enabled", true);            
            FormExtensions.SetControlPropertyThreadSafe(btnApply, "Enabled", true);

            System.Windows.Forms.Application.DoEvents();
        }

        public void PerformProgressStep(string infoText)
        {
            FormExtensions.SetControlPropertyThreadSafe(lblProgressInfo, "Text", infoText);
            
            if (pbMain.Value < pbMain.Maximum)
                FormExtensions.SetControlPropertyThreadSafe(pbMain, "Value", pbMain.Value + 1);                
            
            System.Windows.Forms.Application.DoEvents();
        }

        private void chkUseRubbishPage_CheckedChanged(object sender, EventArgs e)
        {
            SetRubbishControlsAbility();
        }

        private void btnDeleteNotesPages_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show(BibleCommon.Resources.Constants.ConfiguratorQuestionDeleteAllNotesPages, BibleCommon.Resources.Constants.Warning,
                        MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) == System.Windows.Forms.DialogResult.Yes)
            {
                if (SettingsManager.Instance.NotebookId_BibleComments == SettingsManager.Instance.NotebookId_BibleNotesPages
                    || SettingsManager.Instance.StoreNotesPagesInFolder 
                    || MessageBox.Show(BibleCommon.Resources.Constants.ConfiguratorQuestionDeleteAllNotesPagesManually, BibleCommon.Resources.Constants.Warning,
                        MessageBoxButtons.YesNo, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2) == System.Windows.Forms.DialogResult.Yes)
                {
                    using (var manager = new DeleteNotesPagesManager(this))
                    {
                        manager.DeleteNotesPages();
                    }
                }
            }
        }

        private void btnResizeBibleTables_Click(object sender, EventArgs e)
        {
            using (SetWidthForm form = new SetWidthForm())
            {
                if (form.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    using (ResizeBibleManager manager = new ResizeBibleManager(this))
                    {
                        manager.ResizeBiblePages(form.BiblePagesWidth);
                    }
                }
            }
        }

        private void btnBackup_Click(object sender, EventArgs e)
        {
            saveFileDialog.DefaultExt = ".zip";
            saveFileDialog.FileName = string.Format("{0}_backup_{1}", Constants.NewToolsName, DateTime.Now.ToString("yyyy.MM.dd"));

            if (saveFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                BackupManager manager = new BackupManager(this);

                manager.Backup(saveFileDialog.FileName);
            }
        }        

        private void tabPage1_Enter(object sender, EventArgs e)
        {
            if (!_firstShown && !SettingsManager.Instance.CurrentModuleIsCorrect())            
                tbcMain.SelectedTab = tbcMain.TabPages[tpModules.Name];
        }

        private void btnUploadModule_Click(object sender, EventArgs e)
        {
            try
            {
                btnUploadModule.Enabled = false;
                if (openModuleFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    if (Path.GetExtension(openModuleFileDialog.FileName).ToLower() != Constants.FileExtensionIsbt)
                        FormLogger.LogError(BibleCommon.Resources.Constants.ShouldSelectIsbtFile);
                    else
                    {
                    bool moduleWasAdded;
                        bool needToReload = AddNewModule(openModuleFileDialog.FileName, out moduleWasAdded);
                    if (needToReload)
                        ReLoadParameters(true);
                }
            }
            }
            finally
            {
                btnUploadModule.Enabled = true;
            }
        }

        private void ReLoadModulesInfo()
        {
            pnModules.Controls.Clear();
            LoadModulesInfo();
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="needToLoadParameters"></param>
        /// <returns>true если новый модуль стал основным</returns>
        public bool AddNewModule(string filePath, out bool moduleWasAdded)
        {
            var preModuleInfo = ModulesManager.ReadModuleInfo(filePath);
            string moduleName = preModuleInfo.ShortName;            
            
            string destFilePath = Path.Combine(ModulesManager.GetModulesPackagesDirectory(), moduleName + Constants.FileExtensionIsbt);

            moduleWasAdded = true;
            bool canContinue = true;
            if (File.Exists(destFilePath))
            {
                var needToAsk = false;

                ModuleInfo existingModule = null;

                try
                {
                    existingModule = ModulesManager.GetModuleInfo(moduleName);
                    if (existingModule.Version > preModuleInfo.Version) 
                        needToAsk = true;
                }
                catch (InvalidModuleException)
                { }                   

                
                if (needToAsk 
                    && existingModule != null 
                    && MessageBox.Show(string.Format(BibleCommon.Resources.Constants.ModuleWithSameNameAlreadyExists, existingModule.Version, preModuleInfo.Version),
                                                BibleCommon.Resources.Constants.Warning, MessageBoxButtons.YesNo, 
                                                MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1) == System.Windows.Forms.DialogResult.No)
                {
                    canContinue = false;
                    moduleWasAdded = false;
                }
            }

            if (canContinue)
            {
                ModuleInfo module = null;

                try
                {
                    bool currentModuleIsCorrect = SettingsManager.Instance.CurrentModuleIsCorrect();  // а то может быть, что мы загрузили модуль, и он стал корретным, но UI не обновилось

                    module = ModulesManager.UploadModule(filePath, destFilePath, moduleName);

                    bool needToReload = false;

                    if (!currentModuleIsCorrect && module.Type == ModuleType.Bible)
                    {
                        SettingsManager.Instance.ModuleShortName = module.ShortName;
                        needToReload = true;
                    }
                    
                    FormExtensions.Invoke(this, ReLoadModulesInfo);

                    FormLogger.LogMessage(BibleCommon.Resources.Constants.ModuleSuccessfullyUploaded);
                    
                    return needToReload;                    
                }
                catch (InvalidModuleException ex)
                {
                    FormLogger.LogError(ex);                    
                    Thread.Sleep(500);

                    if (module != null)
                        ModulesManager.DeleteModule(module.ShortName);
                    else
                        ModulesManager.DeleteModule(moduleName);

                    moduleWasAdded = false;
                }
            }

            return false;
        }
        

        private bool _wasLoadedModulesInfo = false;

        private void tabPage4_Enter(object sender, EventArgs e)
        {
            if (!_wasLoadedModulesInfo)
            {
                LoadModulesInfo();
            }

            btnSupplementalBibleManagement.Text = BibleCommon.Resources.Constants.SupplementalBibleManagement;
            btnDictionariesManagement.Text = BibleCommon.Resources.Constants.DictionariesManagement;
        }

        private static int GetModuleTypeWeight(ModuleType type)
        {
            switch (type)
            {
                case ModuleType.Bible:
                    return 0;
                default:
                    return 1;
            }
        }


        private bool _lblModulesBibleTitleWasAdded = false;
        private bool _lblModulesDictionariesTitleWasAdded = false;


        private const int MaxPnModulesHeight = 265;
        private void LoadModulesInfo()
        {            

            int top = 10;
            _lblModulesBibleTitleWasAdded = false;
            _lblModulesDictionariesTitleWasAdded = false;
            var allModules = ModulesManager.GetModules(false);
            var modules = new List<ModuleInfo>();
            foreach (var module in allModules.OrderBy(m => GetModuleTypeWeight(m.Type)).ThenBy(m => m.DisplayName))
            {
                try
                {
                    ModulesManager.CheckModule(module);

                    top = SetModulesGroupTitle(module, top);

                    LoadModuleToUI(module, top);
                    modules.Add(module);
                }
                catch (Exception ex)
                {
                    var loadFormTopMost = _loadForm.TopMost;
                    var formTopMost = this.TopMost;

                    string moduleDirectory = ModulesManager.GetModuleDirectory(module.ShortName);
                    _loadForm.TopMost = false;
                    this.TopMost = false;
                    try
                    {
                        FormLogger.LogMessage(string.Format(BibleCommon.Resources.Constants.ModuleUploadError, moduleDirectory, ex.Message));
                        if (DeleteModuleWithConfirm(module.ShortName, true))
                            return;
                    }
                    finally
                    {
                        _loadForm.TopMost = loadFormTopMost;
                        this.TopMost = formTopMost;
                    }
                }

                top += 30;
            }

            if (top > MaxPnModulesHeight)
                top = MaxPnModulesHeight;

            pnModules.Height = top;
            btnUploadModule.Top = top + 50;

            if (modules.Where(m => m.Type == ModuleType.Bible).Count() > 0)
            {                
                btnUploadModule.Left = 31 + pnModules.Left;

                btnSupplementalBibleManagement.Top = btnUploadModule.Top;
                btnSupplementalBibleManagement.Left = btnUploadModule.Right + 6;
                btnSupplementalBibleManagement.Visible = true;

                btnDictionariesManagement.Top = btnUploadModule.Top;
                btnDictionariesManagement.Left = btnSupplementalBibleManagement.Right + 6;
                btnDictionariesManagement.Visible = true;

                lblMustUploadModule.Visible = false;
                lblMustSelectModule.Visible = !SettingsManager.Instance.CurrentModuleIsCorrect();
            }
            else
            {
                if (modules.Count == 0)
                    btnUploadModule.Top = 125;

                lblMustUploadModule.Top = btnUploadModule.Top - 20;                

                btnUploadModule.Left = (this.Width - btnUploadModule.Width) / 2;                

                lblMustUploadModule.Visible = true;
                lblMustSelectModule.Visible = false;
                btnSupplementalBibleManagement.Visible = false;
                btnDictionariesManagement.Visible = false;
            }
            
            _wasLoadedModulesInfo = true;
        }

        private int SetModulesGroupTitle(ModuleInfo module, int top)
        {
            Label lblTitle = null;
            if (module.Type == ModuleType.Bible)
            {
                if (!_lblModulesBibleTitleWasAdded)
                {
                    lblTitle = new Label() { Text = BibleCommon.Resources.Constants.BaseModules, Top = top + 10, Width = 600 };
                    _lblModulesBibleTitleWasAdded = true;
                }
            }
            else
            {
                if (!_lblModulesDictionariesTitleWasAdded)
                {
                    lblTitle = new Label() { Text = BibleCommon.Resources.Constants.AdditionalModules, Top = top + 10, Width = 600 };
                    _lblModulesDictionariesTitleWasAdded = true;
                }
            }

            if (lblTitle != null)
            {
                lblTitle.Font = new Font(lblTitle.Font, FontStyle.Bold);
                pnModules.Controls.Add(lblTitle);
                top += 35;
            }

            return top;
        }

        private void LoadModuleToUI(ModuleInfo moduleInfo, int top)
        {   
            var maximumModuleNameLength = 45;
            Label lblName = new Label();
            if (moduleInfo.DisplayName.Length > maximumModuleNameLength)
                lblName.Text = moduleInfo.DisplayName.Substring(0, maximumModuleNameLength) + "...";
            else
                lblName.Text = moduleInfo.DisplayName;
            lblName.Top = top + 5;
            lblName.Left = 15;
            lblName.Width = 305;
            FormExtensions.SetToolTip(lblName, BibleCommon.Resources.Constants.ModuleDisplayName);
            pnModules.Controls.Add(lblName);

            Label lblVersion = new Label();
            lblVersion.Text = moduleInfo.Version.ToString();
            lblVersion.Top = top + 5;
            lblVersion.Left = 325;
            lblVersion.Width = 30;
            FormExtensions.SetToolTip(lblVersion, BibleCommon.Resources.Constants.ModuleVersion);
            pnModules.Controls.Add(lblVersion);

            if (moduleInfo.Type == ModuleType.Bible)
            {
                CheckBox cbIsActive = new CheckBox();
                cbIsActive.AutoCheck = false;
                cbIsActive.Checked = SettingsManager.Instance.ModuleShortName == moduleInfo.ShortName;
                cbIsActive.Top = top;
                cbIsActive.Left = 355;
                cbIsActive.Width = 20;
                FormExtensions.SetToolTip(cbIsActive, BibleCommon.Resources.Constants.ModuleIsActive);
                pnModules.Controls.Add(cbIsActive);
            }

            
            Button btnInfo = new Button();
            btnInfo.Text = "?";
            btnInfo.Tag = moduleInfo;
            btnInfo.Top = top;
            btnInfo.Left = 375;
            btnInfo.Width = 20;
            btnInfo.Click += new EventHandler(btnModuleInfo_Click);
            FormExtensions.SetToolTip(btnInfo, BibleCommon.Resources.Constants.ModuleInformation);
            pnModules.Controls.Add(btnInfo);
            

            if (moduleInfo.Type == ModuleType.Bible)
            {
                Button btnUseThisModule = new Button();
                btnUseThisModule.Text = GetBtnModuleManagementText(moduleInfo.Type);
                btnUseThisModule.Enabled = moduleInfo.Type == ModuleType.Bible ? SettingsManager.Instance.ModuleShortName != moduleInfo.ShortName : true;
                btnUseThisModule.Tag = moduleInfo;
                btnUseThisModule.Top = top;
                btnUseThisModule.Left = 400;
                btnUseThisModule.Width = 185;
                btnUseThisModule.Click += new EventHandler(btnUseThisModule_Click);
                pnModules.Controls.Add(btnUseThisModule);
            }

            Button btnDel = new Button();
            btnDel.Image = BibleConfigurator.Properties.Resources.del;
            btnDel.Enabled = SettingsManager.Instance.ModuleShortName != moduleInfo.ShortName;            
            btnDel.Tag = moduleInfo.ShortName;
            btnDel.Top = top;            
            btnDel.Left = 590;
            btnDel.Width = btnDel.Height;
            btnDel.Click += new EventHandler(btnDeleteModule_Click);
            FormExtensions.SetToolTip(btnDel, BibleCommon.Resources.Constants.DeleteThisModule);
            pnModules.Controls.Add(btnDel);            
        }       

        private string GetBtnModuleManagementText(ModuleType moduleType)
        {
            switch (moduleType)
            {
                case ModuleType.Bible: 
                    return BibleCommon.Resources.Constants.UseThisModule; 
                case ModuleType.Strong:
                    return BibleCommon.Resources.Constants.SupplementalBibleManagement; 
                case ModuleType.Dictionary:
                    return BibleCommon.Resources.Constants.DictionariesManagement;
                default: 
                    throw new NotSupportedException(moduleType.ToString());
            }
        }

        void btnModuleInfo_Click(object sender, EventArgs e)
        {
            var btn = (Button)sender;
            var moduleInfo = (ModuleInfo)btn.Tag;

            if (moduleInfo.Type == ModuleType.Dictionary)
            {
                MessageBox.Show(!string.IsNullOrEmpty(moduleInfo.Description) ? moduleInfo.Description : moduleInfo.DisplayName, 
                    BibleCommon.Resources.Constants.ModuleInformation, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                using (AboutModuleForm f = new AboutModuleForm(moduleInfo.ShortName, false))
                {
                    f.ShowDialog();
                }
            }
        }

        void btnUseThisModule_Click(object sender, EventArgs e)
        {
            var btn = (Button)sender;
            var moduleInfo = (ModuleInfo)btn.Tag;

            switch (moduleInfo.Type)
            {
                case ModuleType.Bible:
                    bool canContinue = true;

                    if (!string.IsNullOrEmpty(SettingsManager.Instance.NotebookId_Bible) && OneNoteUtils.NotebookExists(ref _oneNoteApp, SettingsManager.Instance.NotebookId_Bible, true)
                        && SettingsManager.Instance.CurrentModuleIsCorrect())
                    {
                        if (moduleInfo.ShortName != _originalModuleShortName)
                        {
                            if (MessageBox.Show(BibleCommon.Resources.Constants.ChangeModuleWarning, BibleCommon.Resources.Constants.Warning,
                                MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == System.Windows.Forms.DialogResult.No)
                            {
                                canContinue = false;
                            }
                            else
                            {
                                _moduleWasChanged = true;
                            }
                        }
                        else
                        {
                            _moduleWasChanged = false;
                        }
                    }

                    if (canContinue)
                    {
                        BibleParallelTranslationManager.RemoveBookAbbreviationsFromMainBible(null, true);
                        SettingsManager.Instance.ModuleShortName = moduleInfo.ShortName;
                        BibleParallelTranslationManager.MergeAllModulesWithMainBible();
                        ReLoadModulesInfo();
                        ReLoadParameters(SettingsManager.Instance.ModuleShortName != _originalModuleShortName);
                        tbcMain.SelectedTab = tbcMain.TabPages[tpNotebooks.Name];
                    }
                    break;
                case ModuleType.Strong:
                    ShowSupplementalBibleManagementForm();
                    break;
                case ModuleType.Dictionary:
                    ShowDictionariesManagementForm();
                    break;
            }
        }        

        internal void ReLoadParameters(bool needToSaveSettings)
        {
            _loadForm.SetDesktopLocation(this.Left - 5, this.Top - 5);
            _loadForm.Show();
            try
            {
                LoadParameters(ModulesManager.GetCurrentModuleInfo(), needToSaveSettings);
            }
            finally
            {
                _loadForm.Hide();
            }
        }

        void btnDeleteModule_Click(object sender, EventArgs e)
        {
            var btn = (Button)sender;
            var moduleName = (string)btn.Tag;

            if (SettingsManager.Instance.SupplementalBibleModules.Any(m => m.ModuleName == moduleName)
                    && !string.IsNullOrEmpty(SettingsManager.Instance.GetValidSupplementalBibleNotebookId(ref _oneNoteApp, true)))
                FormLogger.LogError(BibleCommon.Resources.Constants.ModuleCannotBeDeleted_SupplementalBibleModule);
            else if (SettingsManager.Instance.DictionariesModules.Any(m => m.ModuleName == moduleName) 
                    && !string.IsNullOrEmpty(SettingsManager.Instance.GetValidDictionariesNotebookId(ref _oneNoteApp, true)))
                FormLogger.LogError(BibleCommon.Resources.Constants.ModuleCannotBeDeleted_DictionaryModule);
            else
                DeleteModuleWithConfirm(moduleName, true);
        }

        private bool DeleteModuleWithConfirm(string moduleName, bool toReloadModules)
        {
            using (var form = new MessageForm(BibleCommon.Resources.Constants.DeleteThisModuleQuestion, BibleCommon.Resources.Constants.Warning, MessageBoxButtons.YesNo, MessageBoxIcon.Warning))
            {
                if (form.ShowDialog() == System.Windows.Forms.DialogResult.Yes)
                {
                    ModulesManager.DeleteModule(moduleName);
                    if (toReloadModules)
                ReLoadModulesInfo();
                return true;
            }
            }

            return false;
        }

        private void hlModules_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start(BibleCommon.Resources.Constants.WebSiteUrl + "/modules.htm");
        }

        private void btnSupplementalBibleManagement_Click(object sender, EventArgs e)
        {
            ShowSupplementalBibleManagementForm();
        }

        private void ShowSupplementalBibleManagementForm()
        {
            using (var form = new SupplementalBibleForm(this))
            {
                form.ShowDialog();
            }
        }

        private void btnDictionariesManagement_Click(object sender, EventArgs e)
        {
            ShowDictionariesManagementForm();
        }

        private void ShowDictionariesManagementForm()
        {
            using (var form = new DictionaryModulesForm(this))
            {
                form.ShowDialog();
            }
        }

        private void btnConverter_Click(object sender, EventArgs e)
        {
            var form = new ZefaniaXmlConverterForm(this);

            form.ShowDialog();

            if (form.NeedToCheckModule)
            {
                var moduleCheckForm = new ParallelBibleCheckerForm(this);
                moduleCheckForm.ModuleToCheckName = form.ConvertedModuleShortName;
                moduleCheckForm.AutoStart = true;

                moduleCheckForm.ShowDialog();
            }
        }

        private void btnModuleChecker_Click(object sender, EventArgs e)
        {
            var form = new ParallelBibleCheckerForm(this);

            form.ShowDialog();
        }

        private void chkNotOneNoteControls_CheckedChanged(object sender, EventArgs e)
        {
            if (!((CheckBox)sender).Checked)
            {
                if (SettingsManager.Instance.SupplementalBibleModules != null)
                {
                    var modules = ModulesManager.GetModules(true);
                    if (SettingsManager.Instance.SupplementalBibleModules.Any(moduleInfo =>
                        {
                            var module = modules.FirstOrDefault(m => m.ShortName == moduleInfo.ModuleName);
                            if (module != null)
                                return module.Type == ModuleType.Strong;
                            return false;
                        }))
                    {
                        MessageBox.Show(BibleCommon.Resources.Constants.ChangedNotOneNoteControlsParameter);
                    }
                }
            }
        }

        private void chkUseFolderForBibleNotesPages_CheckedChanged(object sender, EventArgs e)
        {
            SetNotesPagesNotebookControlsAbility();

            SetNotesPageAnalyzeControlsAbility();
            SetRubbishControlsAbility();
        }

        private void btnBibleNotesPagesSetFolder_Click(object sender, EventArgs e)
        {
            if (chkUseFolderForBibleNotesPages.Checked)
            {
                if (notesPagesFolderBrowserDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    tbBibleNotesPagesFolder.Text = notesPagesFolderBrowserDialog.SelectedPath;
                    ttNotesPageFolder.SetToolTip(tbBibleNotesPagesFolder, tbBibleNotesPagesFolder.Text);
                }
            }
        }

        private void tbBibleNotesPagesFolder_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(tbBibleNotesPagesFolder.Text))
                btnBibleNotesPagesSetFolder_Click(this, null);
        }

        private void SetNotesPagesNotebookControlsAbility()
        {
            chkCreateBibleNotesPagesNotebookFromTemplate.Enabled = chkUseFolderForBibleNotesPages.Enabled && !chkUseFolderForBibleNotesPages.Checked;

            cbBibleNotesPagesNotebook.Enabled = (chkUseFolderForBibleNotesPages.Enabled && !chkUseFolderForBibleNotesPages.Checked)
                                                && (chkCreateBibleNotesPagesNotebookFromTemplate.Enabled && !chkCreateBibleNotesPagesNotebookFromTemplate.Checked);            

            tbBibleNotesPagesFolder.Enabled = chkUseFolderForBibleNotesPages.Enabled && chkUseFolderForBibleNotesPages.Checked;
            btnBibleNotesPagesSetFolder.Enabled = chkUseFolderForBibleNotesPages.Enabled && chkUseFolderForBibleNotesPages.Checked;
            
            btnBibleNotesPagesNotebookSetPath.Enabled = chkCreateBibleNotesPagesNotebookFromTemplate.Enabled && chkCreateBibleNotesPagesNotebookFromTemplate.Checked;
        }

        private void SetNotesPageAnalyzeControlsAbility()
        {
            var storeNotesPageInFolder = !chkUseFolderForBibleNotesPages.Enabled || !chkUseFolderForBibleNotesPages.Checked;

            chkExcludedVersesLinking.Enabled = storeNotesPageInFolder && !chkDefaultParameters.Checked;
            chkExpandMultiVersesLinking.Enabled = storeNotesPageInFolder && !chkDefaultParameters.Checked;
            chkUseDifferentPages.Enabled = storeNotesPageInFolder && !chkDefaultParameters.Checked;
            tbNotesPageWidth.Enabled = storeNotesPageInFolder && !chkDefaultParameters.Checked;
        }

        private void SetRubbishControlsAbility()
        {
            var storeNotesPageInFolder = !chkUseFolderForBibleNotesPages.Enabled || !chkUseFolderForBibleNotesPages.Checked;

            chkUseRubbishPage.Enabled = storeNotesPageInFolder && !chkDefaultParameters.Checked;

            var useRubbishPage = chkUseRubbishPage.Enabled && chkUseRubbishPage.Checked;

            tbRubbishNotesPageWidth.Enabled = storeNotesPageInFolder && !chkDefaultParameters.Checked && useRubbishPage;
            tbRubbishNotesPageName.Enabled = storeNotesPageInFolder && !chkDefaultParameters.Checked && useRubbishPage;
            chkRubbishExcludedVersesLinking.Enabled = storeNotesPageInFolder && !chkDefaultParameters.Checked && useRubbishPage;
            chkRubbishExpandMultiVersesLinking.Enabled = storeNotesPageInFolder && !chkDefaultParameters.Checked && useRubbishPage;            
        }
    }
}
