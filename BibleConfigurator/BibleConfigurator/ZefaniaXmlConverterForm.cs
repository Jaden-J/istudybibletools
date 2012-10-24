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

namespace BibleConfigurator
{
    public partial class ZefaniaXmlConverterForm : Form
    {
        public struct LanguageInfo
        {
            public string Locale { get; set; }
            public string BibleBooksInfo_String { get; set; }
        }

        private Dictionary<string, LanguageInfo> LanguageParameters { get; set; }

        public ZefaniaXmlConverterForm()
        {
            InitializeComponent();
        }

        private void ZefaniaXmlConverterForm_Load(object sender, EventArgs e)
        {
            LanguageParameters = new Dictionary<string, LanguageInfo>();
            LanguageParameters.Add("Russian", new LanguageInfo() { Locale = "ru", BibleBooksInfo_String = Properties.Resources.BibleBooskInfo_rst });

            BindControls();            
        }

        private void BindControls()
        {
            cbLanguages.DataSource = LanguageParameters;
            cbLanguages.ValueMember = "Key";
            cbLanguages.DisplayMember = "Key";
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }        

        private void btnOk_Click(object sender, EventArgs e)
        {
            //var converter = new ZefaniaXmlConverter("ibs", "Современный перевод (Всемирный Библейский Переводческий Центр)", 
            //    @"C:\Users\lux_demko\Desktop\temp\Dropbox\Holy Bible\ForGenerating\ibs\bible.xml",
            //    Utils.LoadFromXmlString<BibleBooksInfo>(Properties.Resources.BibleBooskInfo_rst), @"c:\temp\ibsZefania", "ru",
            //    PredefinedNotebooksInfo.Russian, Utils.LoadFromXmlString<BibleTranslationDifferences>(Properties.Resources.BibleTranslationDifferences_rst),  // вот эти тоже часто надо менять                
            //    "{0} глава. {1}",
            //    PredefinedSectionsInfo.None, false, null, null,
            //    //PredefinedSectionsInfo.RSTStrong, true, "Стронга", 14700,   // параметры для стронга
            //    "2.0", true,
            //    ZefaniaXmlConverter.ReadParameters.None);  // и про эту не забыть

            //converter.Convert();
        }      
    }
}
