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

namespace BibleCommon.UI.Forms
{
    public partial class ErrorsForm : Form
    {
        public List<string> Errors { get; set; }

        public ErrorsForm(List<string> errors)
        {
            Errors = errors;

            InitializeComponent();            
        }

        private void Errors_Load(object sender, EventArgs e)
        {
            foreach(string error in Errors)
            {
                lbErrors.Items.Add(error);

                int width = Convert.ToInt32(error.Length * 5.75);
                if (width > lbErrors.HorizontalExtent)
                    lbErrors.HorizontalExtent = width;          
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
                        foreach (var error in Errors)
                        {
                            sw.WriteLine(error);
                        }
                        sw.Flush();
                    }
                }                
            }
        }
    }
}
