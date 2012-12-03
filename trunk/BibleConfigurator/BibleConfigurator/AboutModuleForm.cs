using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using BibleCommon.Services;
using BibleCommon.Common;
using System.IO;
using BibleCommon.Helpers;

namespace BibleConfigurator
{
    public partial class AboutModuleForm : Form
    {
        public string ModuleName { get; set; }

        public AboutModuleForm(string moduleName, bool topMost)
        {
            this.SetFormUICulture();

            InitializeComponent();

            this.ModuleName = moduleName;
            this.TopMost = topMost;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void AboutModule_Load(object sender, EventArgs e)
        {
            try
            {
                ModuleInfo module = ModulesManager.GetModuleInfo(ModuleName);

                this.Text = lblTitle.Text = string.Format("{0} ({1} {2})", module.DisplayName, module.BibleStructure.BibleBooks.Count, BibleCommon.Resources.Constants.Books);
                lblLocation.Text = ModulesManager.GetModuleDirectory(ModuleName);

                var sb = new StringBuilder(
@"<html>
    <body>
        <div>");
                sb.AppendFormat(
@"
            <table style='font-family: @{0};font-size:small'>", BibleCommon.Consts.Constants.UnicodeFontName);
                foreach (var book in module.BibleStructure.BibleBooks)
                {
                    sb.Append(
@"
                <tr>");
                    sb.AppendFormat(
@"
                    <td>
                        <span style='font-weight: bold;white-space:nowrap;'>{0}</span>:
                    </td>", book.Name);
                    sb.AppendFormat(
@"
                    <td style='padding-left: 10px;'>
                        <span style='white-space:nowrap;'>{0}</span>
                    </td>", string.Join(",&nbsp;", book.Abbreviations.Select(abbr => string.Format("'{0}'", abbr.Value)).ToArray()));
                    sb.Append(
@"
                </tr>");
                }
                sb.Append(
@"
            </table>
        </div>
    </body>
</html>");

                wbBooks.DocumentText = sb.ToString();              
            }
            catch (Exception ex)
            {
                FormLogger.LogError(ex);
            }
        }      

        private bool _wasShown = false;
        private void AboutModuleForm_Shown(object sender, EventArgs e)
        {   
            if (!_wasShown)
            {
                this.SetFocus();
                _wasShown = true;
            }
        }
    }
}
