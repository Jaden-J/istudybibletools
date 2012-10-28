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
            MessageBox.Show(ex.Message, BibleCommon.Resources.Constants.Warning, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            WasErrorLogged = true;
        }

        public static void LogError(string message)
        {
            BibleCommon.Services.Logger.LogError(message);
            MessageBox.Show(message, BibleCommon.Resources.Constants.Warning, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            WasErrorLogged = true;
        }

        public static void LogMessage(string message)
        {
            BibleCommon.Services.Logger.LogMessage(message);
            MessageBox.Show(message, string.Empty, MessageBoxButtons.OK, MessageBoxIcon.Information);            
        }        
    }
}
