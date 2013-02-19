﻿using System;
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
        private IProtocolHandler[] _handlers = new IProtocolHandler[] { new NavigateToStrongHandler(), new FindVersesWithStrongNumberHandler() };

        public CommandForm()
        {   
            InitializeComponent();            

            //Command-Line            
            var args = Environment.GetCommandLineArgs();
            if (args.Length > 0)
                ProcessCommandLine(args);
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
            foreach (var handler in _handlers)
            {
                if (handler.IsProtocolCommand(args))
                {
                    handler.ExecuteCommand(args);
                    break;
                }
            }
        }
    }
}