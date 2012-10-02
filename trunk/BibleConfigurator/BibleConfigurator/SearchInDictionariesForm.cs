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

namespace BibleConfigurator
{
    public partial class SearchInDictionariesForm : Form
    {
        public class TermDictionariesInfo
        {
            public Dictionary<string, string> ModulesWithTerm { get; set; }

            public TermDictionariesInfo(string firstModuleShortName, string firstModuleName)
            {
                ModulesWithTerm = new Dictionary<string,string>();
                ModulesWithTerm.Add(firstModuleShortName, firstModuleName);
            }

            public void AddTermModue(string moduleShortName, string moduleName)
            {
                if (!ModulesWithTerm.ContainsKey(moduleShortName))
                    ModulesWithTerm.Add(moduleShortName, moduleName);
            }
        }
        
        private Dictionary<ModuleInfo, ModuleDictionaryInfo> _modules;

        public Dictionary<string, TermDictionariesInfo> TermsInDictionariesInfo;

        public Dictionary<ModuleInfo, ModuleDictionaryInfo> Modules
        {
            get
            {
                if (_modules == null)
                {
                    try
                    {
                        _modules = new Dictionary<ModuleInfo, ModuleDictionaryInfo>();
                        TermsInDictionariesInfo = new Dictionary<string, TermDictionariesInfo>();

                        foreach (var module in ModulesManager.GetModules())
                        {
                            if (module.Type == ModuleType.Dictionary)
                            {
                                var dictionaryModuleInfo = ModulesManager.GetModuleDictionaryInfo(module.ShortName);
                                _modules.Add(module, dictionaryModuleInfo);

                                foreach (var term in dictionaryModuleInfo.TermSet.Terms)
                                {
                                    if (!TermsInDictionariesInfo.ContainsKey(term))
                                        TermsInDictionariesInfo.Add(term, new TermDictionariesInfo(module.ShortName, module.Name));
                                    else
                                        TermsInDictionariesInfo[term].AddTermModue(module.ShortName, module.Name);
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.LogError(ex);
                        throw;
                    }
                }

                return _modules;
            }
        }

        public SearchInDictionariesForm()
        {
            InitializeComponent();
        }

        private void SearchInDictionariesForm_Load(object sender, EventArgs e)
        {
            var allTerms = Modules.Values.SelectMany(md => md.TermSet.Terms).OrderBy(t => t).ToArray();
            cbAllTerms.Items.AddRange(allTerms);
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            
        }

        private void cbAllTerms_SelectedIndexChanged(object sender, EventArgs e)
        {
            string term = (string)cbAllTerms.SelectedItem;
            if (TermsInDictionariesInfo.ContainsKey(term))
            {
                cbFoundInDictionaries.Items.Clear();            
                cbFoundInDictionaries.Items.AddRange(TermsInDictionariesInfo[term].ModulesWithTerm.Values.ToArray());
                cbFoundInDictionaries.SelectedIndex = 0;
            }
        }        
    }
}
