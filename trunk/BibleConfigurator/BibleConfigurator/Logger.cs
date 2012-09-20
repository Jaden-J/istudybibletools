using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace BibleConfigurator
{
    public static class FormLogger
    {
        public static bool WasErrorLogged = false;

        public static void Initialize()
        {
            WasErrorLogged = false;
        }

        public static void LogError(string message)
        {
            MessageBox.Show(message);
            WasErrorLogged = true;
        }

        public static void LogMessage(string message)
        {
            MessageBox.Show(message);            
        }
    }
}
