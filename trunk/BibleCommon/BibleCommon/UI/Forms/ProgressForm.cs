using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace BibleCommon.UI.Forms
{
    public partial class ProgressForm : Form
    {
        private Action<ProgressForm> _progressAction;             

        public int ProgressStep
        {
            get
            {
                return pbMain.Step;
            }
            set
            {
                pbMain.Step = value;
            }
        }

        public ProgressForm()
        {
            InitializeComponent();
        }

        public ProgressForm(string title, Action<ProgressForm> progressAction)
            : this()
        {
            _progressAction = progressAction;
            Text = title;
        }

        public void PerformStep(string message)
        {
            pbMain.PerformStep();
            lblMessage.Text = message;
            Application.DoEvents();
        }                

        private bool _firstTimeShown = true;
        private void ProgressForm_Shown(object sender, EventArgs e)
        {
            if (_firstTimeShown)
            {
                _firstTimeShown = false;

                TopMost = true;
                Focus();
                Application.DoEvents();
                TopMost = false;

                if (_progressAction != null)
                    _progressAction(this);

                Close();
            }
        }

        public void ShowDialog(int maximumProgressValue)
        {
            pbMain.Maximum = maximumProgressValue;
            base.ShowDialog();            
        }
    }
}
