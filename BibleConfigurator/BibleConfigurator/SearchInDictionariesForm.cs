using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using BibleCommon.Common;
using BibleCommon.Services;
using BibleCommon.Helpers;
using System.Threading;

namespace BibleConfigurator
{
    public partial class SearchInDictionariesForm : Form
    {
        private Microsoft.Office.Interop.OneNote.Application _oneNoteApp;

        private Dictionary<string, ModuleDictionaryInfo> _modulesTermSets;
        private Dictionary<string, ModuleInfo> _modules;
        private Dictionary<string, TermInModules> _termsInModules;


        private static string _lastSearchedTerm = string.Empty;

        protected string DictionariesNotebookId { get; set; }

        public Dictionary<string, ModuleDictionaryInfo> ModulesTermSets
        {
            get
            {
                if (_modulesTermSets == null)                
                    LoadData();                  

                return _modulesTermSets;
            }
        }

        public Dictionary<string, ModuleInfo> Modules
        {
            get
            {
                if (_modules == null)
                    LoadData();

                return _modules;
            }
        }

        public Dictionary<string, TermInModules> TermsInModules
        {
            get
            {
                if (_termsInModules == null)
                    LoadData();

                return _termsInModules;
            }
        }

        public class TermInModules : List<string>
        {
            public string Term { get; set; }
        }

        private bool LoadData()
        {
            try
            {                
                DictionariesNotebookId = SettingsManager.Instance.GetValidDictionariesNotebookId(_oneNoteApp);

                if (string.IsNullOrEmpty(DictionariesNotebookId))
                {
                    SettingsManager.Instance.ReLoadSettings();
                    DictionariesNotebookId = SettingsManager.Instance.GetValidDictionariesNotebookId(_oneNoteApp, true);
                }

                if (string.IsNullOrEmpty(DictionariesNotebookId))
                {
                    this.Visible = false;
                    this.SetFocus();                    
                    FormLogger.LogError(BibleCommon.Resources.Constants.DictionariesNotInstalled);                    
                    Close();
                    return false;
                }

                _modulesTermSets = new Dictionary<string, ModuleDictionaryInfo>();
                _modules = new Dictionary<string, ModuleInfo>();
                _termsInModules = new Dictionary<string, TermInModules>(StringComparer.InvariantCultureIgnoreCase);

                foreach (var module in SettingsManager.Instance.DictionariesModules)
                {
                    var moduleInfo = ModulesManager.GetModuleInfo(module.ModuleName);
                    if (moduleInfo.Type == ModuleType.Dictionary)
                    {
                        var dictionaryModuleInfo = OneNoteProxy.Instance.GetModuleDictionary(moduleInfo.ShortName);
                        _modulesTermSets.Add(moduleInfo.ShortName, dictionaryModuleInfo);
                        _modules.Add(moduleInfo.DisplayName, moduleInfo);

                        foreach (var term in dictionaryModuleInfo.TermSet.Terms)
                        {
                            if (!_termsInModules.ContainsKey(term))
                                _termsInModules.Add(term, new TermInModules() { Term = term });

                            _termsInModules[term].Add(moduleInfo.DisplayName);
                        }
                    }
                }             

                return true;
            }
            catch (Exception ex)
            {
                Logger.LogError(ex);
                BibleCommon.Services.Logger.LogError(ex);
                throw;
            }
        }

        public SearchInDictionariesForm()
        {
            this.SetFormUICulture();
            InitializeComponent();
            _oneNoteApp = new Microsoft.Office.Interop.OneNote.Application();

            this.Text = BibleCommon.Resources.Constants.SearchInDictionaries;
            this.btnCancel.Text = BibleCommon.Resources.Constants.Close;
        }

        private void SearchInDictionariesForm_Load(object sender, EventArgs e)
        {
            try
            {
                lblFoundInDictionaries.Text = string.Empty;
                if (LoadData())
                {
                    LoadTerms(null);

                    cbDictionaries.Items.Add(BibleCommon.Resources.Constants.AllDictionaries);
                    foreach (var dName in Modules.Keys)
                        cbDictionaries.Items.Add(dName);
                    cbDictionaries.SelectedIndex = 0;

                    if (!string.IsNullOrEmpty(_lastSearchedTerm))
                        cbTerms.SelectedItem = _lastSearchedTerm;

                    cbTerms.Select();
                }
            }
            catch (Exception ex)
            {
                FormLogger.LogError(ex);
            }
        }        

        private void LoadTerms(string moduleShortName)
        {            
            var terms = !string.IsNullOrEmpty(moduleShortName) 
                ? ModulesTermSets[moduleShortName].TermSet.Terms 
                : TermsInModules.Keys.ToList();

            var source = terms.OrderBy(t => t).ToArray();

            cbTerms.DataSource = source;

            //var autoCompleteSource = new AutoCompleteStringCollection();
            //autoCompleteSource.AddRange(source);
            //cbTerms.AutoCompleteCustomSource = autoCompleteSource;
            //cbTerms.AutoCompleteSource = AutoCompleteSource.CustomSource;
            //cbTerms.AutoCompleteMode = AutoCompleteMode.Suggest;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            var term = (string)cbTerms.SelectedItem;

            if (string.IsNullOrEmpty(term))
            {
                if (cbTerms.Items.Contains(cbTerms.Text))
                    term = cbTerms.Text;
            }

            if (!string.IsNullOrEmpty(term))
            {
                btnOk.Enabled = false;
                StartTermSearhing(term, (string)cbFoundInDictionaries.SelectedItem);
                Close();
            }
            else
                FormLogger.LogMessage(BibleCommon.Resources.Constants.SelectWord);
        }

        private void SearchInDictionariesForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            _oneNoteApp = null;
        }

        private void cbDictionaries_SelectedIndexChanged(object sender, EventArgs e)
        {
            var selectedDictionary = (string)cbDictionaries.SelectedItem;

            if (selectedDictionary == BibleCommon.Resources.Constants.AllDictionaries)
                LoadTerms(null);
            else
            {
                LoadTerms(Modules[selectedDictionary].ShortName);   
            }
        }

        private void StartTermSearhing(string term, string dictionaryName)
        {
            try
            {
                _lastSearchedTerm = term;
                var moduleShortName = _modules[dictionaryName].ShortName;

                if (!DictionaryTermsCacheManager.CacheIsActive(moduleShortName))
                    throw new Exception(BibleCommon.Resources.Constants.DictionaryCacheFileNotFound);

                var link = OneNoteProxy.Instance.GetDictionaryTermLink(term.ToLower(), moduleShortName);
                _oneNoteApp.NavigateTo(link.PageId, link.ObjectId);
            }
            catch (Exception ex)
            {
                FormLogger.LogError(ex);
            }
        }


        bool _firstTimeEnterWasPressed = false;
        private void SearchInDictionariesForm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (_firstTimeEnterWasPressed)
                {
                    btnOk_Click(this, null);
                    _firstTimeEnterWasPressed = false;
                }
                else
                    _firstTimeEnterWasPressed = true;
            }
            else if (e.KeyCode == Keys.Escape)
                Close();
        }
      

        private bool _wasShown = false;
        private void SearchInDictionariesForm_Shown(object sender, EventArgs e)
        {
            if (!_wasShown)
            {
                this.SetFocus();
                _wasShown = true;
            }
        }

        private DateTime _lastMouseClickedTime = DateTime.Now;
        private string _lastSelectedItem = null;
        private void cbTerms_MouseClick(object sender, MouseEventArgs e)
        {   
            if (DateTime.Now.CompareTo(_lastMouseClickedTime.AddMilliseconds(SystemInformation.DoubleClickTime)) < 1)
            {                
                if (_lastSelectedItem == (string)cbTerms.SelectedItem)
                    DoubleClicked();
            }

            _lastMouseClickedTime = DateTime.Now;
            _lastSelectedItem = (string)cbTerms.SelectedItem;
        }

        private void DoubleClicked()
        {            
            btnOk_Click(this, null);
        }

        private void cbTerms_SelectedIndexChanged(object sender, EventArgs e)
        {
            var term = (string)cbTerms.SelectedItem;
            if (!string.IsNullOrEmpty(term))
            {
                cbFoundInDictionaries.Items.Clear();
                foreach(var name in _termsInModules[term])
                {
                    if (cbDictionaries.SelectedIndex == 0 || (string)cbDictionaries.SelectedItem == name)
                        cbFoundInDictionaries.Items.Add(_modules[name].DisplayName);
                }

                if (cbFoundInDictionaries.Items.Count == 1)
                {
                    lblFoundInDictionaries.Text = BibleCommon.Resources.Constants.FoundInOneDictionary;
                    //cbFoundInDictionaries.Enabled = false;
                }
                else
                {
                    lblFoundInDictionaries.Text = BibleCommon.Resources.Constants.FoundInSeveralDictionaries;
                    //cbFoundInDictionaries.Enabled = true;
                }

                if (cbFoundInDictionaries.Items.Count > 0)
                    cbFoundInDictionaries.SelectedIndex = 0;
            }
        }            
    }
}
