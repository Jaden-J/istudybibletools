using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BibleCommon.Contracts;
using System.Windows.Forms;
using BibleCommon.Common;

namespace BibleConfigurator
{
    public class LongProcessLogger: ICustomLogger, IDisposable
    {
        private MainForm _form;
        public string Preffix { get; set; }
        public bool AbortedByUsers { get; set; }

        public LongProcessLogger(MainForm form)        
        {
            _form = form;
        }

        public void LogMessage(string message, params object[] args)
        {
            LogMessageInternal(message, args);         
        }

        public void LogWarning(string message, params object[] args)
        {
            LogMessageInternal(string.Format("Warning: {0}", message), args);
            
        }

        public void LogException(string message, params object[] args)
        {
            LogMessageInternal(string.Format("Error: {0}", message), args);            
        }

        private void LogMessageInternal(string message, params object[] args)        
        {
            if (AbortedByUsers)
                throw new ProcessAbortedByUserException();

            _form.PerformProgressStep(FormatMessage(message, args));
            Application.DoEvents();
        }

        private string FormatMessage(string message, params object[] args)
        {
            message = string.Format(message, args);

            if (!string.IsNullOrEmpty(Preffix))
                message = string.Format("{0}{1}", Preffix, message);

            return message;
        }

        public void Dispose()
        {
            _form = null;
        }
    }
}
