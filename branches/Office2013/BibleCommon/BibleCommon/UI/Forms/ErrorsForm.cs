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
        public class ErrorsList : List<string>
        {
            public string ErrorsDecription { get; set; }

            public ErrorsList(IEnumerable<string> collection)
                : base(collection)
            {
            }
        }

        public List<ErrorsList> AllErrors { get; set; }

        public ErrorsForm()
        {
            AllErrors = new List<ErrorsList>();

            InitializeComponent();            
        }

        public ErrorsForm(List<string> errors)
            : this()
        {
            AllErrors.Add(new ErrorsList(errors));
        }

        public void ClearErrors()
        {
            AllErrors.Clear();
            lbErrors.Items.Clear();
        }

        private void Errors_Load(object sender, EventArgs e)
        {
            try
            {
                if (AllErrors.All(errors => errors.Count == 0))
                    Close();

                FormExtensions.SetFocus(this);

                using (Graphics g = lbErrors.CreateGraphics())
                {
                    foreach (var errors in AllErrors)
                    {
                        if (!string.IsNullOrEmpty(errors.ErrorsDecription))
                            lbErrors.Items.Add(errors.ErrorsDecription);

                        int index = 1;

                        foreach (string error in errors)
                        {
                            lbErrors.Items.Add(string.Format("{0}. {1}", index++, error));

                            //int width = Convert.ToInt32(error.Length * 5.75);
                            int width = (int)g.MeasureString(error, lbErrors.Font).Width;
                            if (width > lbErrors.HorizontalExtent)
                                lbErrors.HorizontalExtent = width;
                        }
                        lbErrors.Items.Add(string.Empty);
                    }
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
                        foreach (var errors in AllErrors)
                        {
                            if (!string.IsNullOrEmpty(errors.ErrorsDecription))
                                sw.WriteLine(errors.ErrorsDecription);

                            int index = 1;
                            foreach (var error in errors)
                            {
                                sw.WriteLine(string.Format("{0}. {1}", index++, error));
                            }
                            sw.WriteLine(string.Empty);                            
                        }
                        sw.Flush();
                    }
                }

                MessageBox.Show("Successfully.");
            }
        }
    }
}
