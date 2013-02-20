using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using BibleCommon.Contracts;
using BibleCommon.Handlers;

namespace ISBTCommandHandler
{
    public partial class CommandForm : Form
    {
        private IProtocolHandler[] _handlers = new IProtocolHandler[] 
                                                    { 
                                                        new QuickAnalyzeHandler(), 
                                                        new NavigateToStrongHandler(), 
                                                        new FindVersesWithStrongNumberHandler(),
                                                        new RefreshCacheHandler()
                                                    };

        public CommandForm()
        {   
            InitializeComponent();                        
        }

        protected override CreateParams CreateParams
        {
            get
            {
                var cp = base.CreateParams;
                cp.ExStyle |= 0x80;  // Turn on WS_EX_TOOLWINDOW
                return cp;
            }
        }       

        internal void ProcessCommandLine(params string[] args)
        {   
            if (args.Length > 1)
                args = args.ToList().Skip(1).ToArray();            

            foreach (var handler in _handlers)
            {
                if (handler.IsProtocolCommand(args))
                {
                    handler.ExecuteCommand(args);
                    break;
                }
            }
        }

        private bool _firstShown = true;      

        private void CommandForm_Enter(object sender, EventArgs e)
        {
            if (_firstShown)
            {
                _firstShown = false;

                var args = Environment.GetCommandLineArgs();
                if (args.Length > 0)
                    ProcessCommandLine(args);
            }
        }
    }
}
