using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using BibleCommon.Services;
using System.IO;
using BibleCommon.Helpers;

namespace BibleCommon.UI.Forms
{
    public partial class ErrorsForm : Form
    {
        public List<string> Errors { get; set; }
        public string ErrorsDecription { get; set; }

        public ErrorsForm(List<string> errors)
        {   
            Errors = errors;            

            InitializeComponent();            
        }

        private void Errors_Load(object sender, EventArgs e)
        {
            try
            {
                if (Errors.Count == 0)
                    Close();

                FormExtensions.SetFocus(this);

                if (!string.IsNullOrEmpty(ErrorsDecription))
                    lbErrors.Items.Add(ErrorsDecription);

                int index = 1;
                foreach (string error in Errors)
                {
                    lbErrors.Items.Add(string.Format("{0}. {1}", index++, error));

                    int width = Convert.ToInt32(error.Length * 5.75);
                    if (width > lbErrors.HorizontalExtent)
                        lbErrors.HorizontalExtent = width;
                }
            }
            catch (Exception ex)
            {
                FormLogger.LogError(ex);
            }
        }

        private void btnSaveToFile_Click(object sender, EventArgs e)
        {
            if (saveFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                using (FileStream fs = new FileStream(saveFileDialog.FileName, FileMode.Create))
                {
                    using (StreamWriter sw = new StreamWriter(fs))
                    {
                        if (!string.IsNullOrEmpty(ErrorsDecription))
                            sw.WriteLine(ErrorsDecription);

                        int index = 1;
                        foreach (var error in Errors)
                        {
                            sw.WriteLine(string.Format("{0}. {1}", index++, error));
                        }
                        sw.Flush();
                    }
                }

                MessageBox.Show("Successfully.");
            }
        }
    }
}
