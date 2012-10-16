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

namespace BibleConfigurator
{
    public partial class SearchInDictionariesForm : Form
    {
        private Microsoft.Office.Interop.OneNote.Application _oneNoteApp;

        private Dictionary<string, ModuleDictionaryInfo> _modulesTermSets;
        private Dictionary<string, ModuleInfo> _modules;
        private Dictionary<string, List<string>> _termsInModules;

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

                foreach (var module in ModulesManager.GetModules())
                {
                    if (module.Type == ModuleType.Dictionary)
                    {
                        var dictionaryModuleInfo = OneNoteProxy.Instance.GetModuleDictionary(module.ShortName);
                        _modulesTermSets.Add(module.ShortName, dictionaryModuleInfo);
                        _modules.Add(module.Name, module);

                        _termsInModules = new Dictionary<string, List<string>>();

                        foreach (var term in dictionaryModuleInfo.TermSet.Terms)
                        {
                            if (!_termsInModules.ContainsKey(term))
                                _termsInModules.Add(term, new List<string>());

                            _termsInModules[term].Add(module.ShortName);
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
        }

        private void SearchInDictionariesForm_Load(object sender, EventArgs e)
        {
            try
            {
                if (LoadData())
                {
                    LoadDictionary(null);

                    cbDictionaries.Items.Add(BibleCommon.Resources.Constants.AllDictionaries);
                    foreach (var dName in Modules.Keys)
                        cbDictionaries.Items.Add(dName);
                    cbDictionaries.SelectedIndex = 0;

                    cbTerms.Select();
                }
            }
            catch (Exception ex)
            {
                FormLogger.LogError(ex);
            }
        }        

        private void LoadDictionary(string moduleShortName)
        {            
            var terms = !string.IsNullOrEmpty(moduleShortName) 
                ? ModulesTermSets[moduleShortName].TermSet.Terms 
                : ModulesTermSets.Values.SelectMany(md => md.TermSet.Terms).Distinct();
            cbTerms.DataSource = terms.OrderBy(t => t).ToArray();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            var term = (string)cbTerms.SelectedItem;
            if (!string.IsNullOrEmpty(term))
            {
                btnOk.Enabled = false;
                StartTermSearhing(term);
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
                LoadDictionary(null);
            else
            {
                LoadDictionary(Modules[selectedDictionary].ShortName);   
            }
        }

        private void StartTermSearhing(string term)
        {
            string xml;
            _oneNoteApp.FindPages(DictionariesNotebookId, term, out xml, true, true);
        }
      

        private void SearchInDictionariesForm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                btnOk_Click(this, null);
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
            здесь: надо запускать асинхронный процесс поиска и обновлять lbl и cb в другом потоке. Либо делать это всё синхронно, просто по таймеру.
        } 
    }
}
