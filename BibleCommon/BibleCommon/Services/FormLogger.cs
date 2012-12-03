using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace BibleCommon.Services
{
    public static class FormLogger
    {
        public static bool WasErrorLogged = false;

        public static void Initialize()
        {
            WasErrorLogged = false;
        }

        public static void LogError(Exception ex)
        {
            BibleCommon.Services.Logger.LogError(ex);

            using (var form = new BibleCommon.UI.Forms.MessageForm(ex.Message, BibleCommon.Resources.Constants.Warning, MessageBoxButtons.OK, MessageBoxIcon.Warning))
            {
                form.ShowDialog();
            }            

            WasErrorLogged = true;
        }

        public static void LogError(string message, params object[] args)
        {
            if (args != null && args.Length > 0)
                message = string.Format(message, args);

            BibleCommon.Services.Logger.LogError(message);

            using (var form = new BibleCommon.UI.Forms.MessageForm(message, BibleCommon.Resources.Constants.Warning, MessageBoxButtons.OK, MessageBoxIcon.Warning))
            {
                form.ShowDialog();
            }            

            WasErrorLogged = true;
        }

        public static void LogMessage(string message, params object[] args)
        {
            if (args != null && args.Length > 0)
                message = string.Format(message, args);

            BibleCommon.Services.Logger.LogMessage(message);

            using (var form = new BibleCommon.UI.Forms.MessageForm(message, string.Empty, MessageBoxButtons.OK, MessageBoxIcon.Information))
            {
                form.ShowDialog();
            }                
        }        
    }
}
