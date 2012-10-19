using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.OneNote;
using System.Diagnostics;

namespace CommandHandler
{
    public class NavigateToOneNoteHandler
    {
        private static string _protocolName = "isbtOpen";

        public string ProtocolName
        {
            get { return _protocolName; }
        }

        public static string GetCommandUrl(string pageId, string objectId)
        {
            return string.Format("{0}:{1};{2}", _protocolName, pageId, objectId);
        }

        public string GetCommandUrl(string args)
        {
            return string.Format("{0}:{1}", ProtocolName, args);
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
            var pageId = verseArgs[0];
            var objectId = verseArgs[1];

            Microsoft.Office.Interop.OneNote.Application oneNoteApp = new Microsoft.Office.Interop.OneNote.Application();
            try
            {
                oneNoteApp.NavigateTo(pageId, objectId);
            }
            finally
            {
                oneNoteApp = null;
            }
        }
    }
}
