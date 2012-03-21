using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace BibleNoteLinkerEx
{
    public partial class SelectNoteBooks : Form
    {
        public SelectNoteBooks()
        {
            InitializeComponent();
        }

        private void SelectNoteBooks_Load(object sender, EventArgs e)
        {
            Dictionary<string, string> allNotebooks = GetAllNotebooks();
        }
    }
}
