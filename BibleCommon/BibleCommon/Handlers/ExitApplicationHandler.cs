using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BibleCommon.Contracts;
using System.Windows.Forms;

namespace BibleCommon.Handlers
{
    public class ExitApplicationHandler : IProtocolHandler
    {
        public string ProtocolName
        {
            get { return "isbtExitApplication:"; }
        }

        public string GetCommandUrl(string args)
        {
            return string.Format("{0}{1}", ProtocolName, "exit");
        }

        public bool IsProtocolCommand(params string[] args)
        {
            return args.Length > 0 && args[0].StartsWith(ProtocolName, StringComparison.OrdinalIgnoreCase);
        }

        public void ExecuteCommand(params string[] args)
        {
            Application.Exit();
        }
    }
}
