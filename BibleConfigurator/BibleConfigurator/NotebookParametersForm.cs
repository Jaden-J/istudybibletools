using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace BibleConfigurator
{
    public partial class NotebookParametersForm : Form
    {
        Microsoft.Office.Interop.OneNote.Application OneNoteApp = new Microsoft.Office.Interop.OneNote.Application();

        public NotebookParametersForm()
        {
            InitializeComponent();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {

        }

        private void NotebookParametersForm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                Close();
        }

        private void NotebookParametersForm_Load(object sender, EventArgs e)
        {
            
        }

       
    }
}
