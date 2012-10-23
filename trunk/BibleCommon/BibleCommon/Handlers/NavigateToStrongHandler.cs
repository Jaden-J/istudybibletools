using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.OneNote;
using System.Diagnostics;
using BibleCommon.Contracts;
using BibleCommon.Services;

namespace BibleCommon.Handlers
{
    public class NavigateToStrongHandler : IProtocolHandler
    {
        private static string _protocolName = "isbtStrongOpen";

        public string ProtocolName
        {
            get { return _protocolName; }
        }

        public static string GetCommandUrlStatic(string strongNumber, string moduleShortName)
        {
            return string.Format("{0}:{1};{2}", _protocolName, strongNumber, moduleShortName);
        }      

        public bool IsProtocolCommand(string[] args)
        {
            return args.Length > 0 && args[0].StartsWith(ProtocolName, StringComparison.OrdinalIgnoreCase);
        }

        public void ExecuteCommand(string[] args)
        {
            if (args.Length == 0)
                throw new ArgumentNullException("args");

            var verseArgs = Uri.UnescapeDataString(args[0]
                                .Split(new char[] { ':' })[1])
                                .Split(new char[] { ';' });

            var strongNumber = verseArgs[0];
            var moduleShortName = verseArgs[1];
            var strongTermLink = OneNoteProxy.Instance.GetDictionaryTermLink(strongNumber, moduleShortName);

            Application oneNoteApp = new Application();
            try
            {
                oneNoteApp.NavigateTo(strongTermLink.PageId, strongTermLink.ObjectId);
            }
            finally
            {
                oneNoteApp = null;
            }
        }

        string IProtocolHandler.GetCommandUrl(string args)
        {
            return string.Format("{0}:{1}", ProtocolName, args);
        }      
    }
}
