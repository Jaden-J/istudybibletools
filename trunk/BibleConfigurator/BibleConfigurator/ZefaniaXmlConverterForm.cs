﻿using System;
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
        public ZefaniaXmlConverterForm()
        {
            InitializeComponent();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            var converter = new ZefaniaXmlConverter("Русский синодальный перевод", "rst", @"C:\Users\lux_demko\Desktop\temp\Dropbox\Holy Bible\ForGenerating\RSTStrong\bible.xml",
                Utils.LoadFromXmlString<BibleBooksInfo>(Properties.Resources.BibleBooskInfo_rst), @"c:\temp\rstZefania", "ru",
                PredefinedNotebooksInfo.Russian, Utils.LoadFromXmlString<BibleTranslationDifferences>(Properties.Resources.BibleTranslationDifferences_rst),  // вот эти тоже часто надо менять                
                "{0} глава. {1}",
                PredefinedSectionsInfo.None, false, null, null,
                //PredefinedSectionsInfo.RSTStrong, true, "Стронга", 14700,   // параметры для стронга
                "2.0", false,
                ZefaniaXmlConverter.ReadParameters.RemoveStrongs);  // и про эту не забыть

            converter.Convert();
        }
    }
}