using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using BibleCommon.Services;

namespace BibleNoteLinker
{
    public partial class ErrorsForm : Form
    {
        public ErrorsForm()
        {
            InitializeComponent();
        }

        private void Errors_Load(object sender, EventArgs e)
        {
            foreach(string error in Logger.Errors)
            {
                lbErrors.Items.Add(error);

                int width = Convert.ToInt32(error.Length * 5.75);
                if (width > lbErrors.HorizontalExtent)
                    lbErrors.HorizontalExtent = width;          
            }           
            
        }
    }
}
