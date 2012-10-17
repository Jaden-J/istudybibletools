using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BibleCommon.Contracts;
using Microsoft.Office.Interop.OneNote;
using BibleCommon.Services;
using System.Runtime.InteropServices;

namespace BibleCommon.Handlers
{
    public class FindVersesWithStrongNumberHandler : IProtocolHandler
    {
        public string ProtocolName
        {
            get { return "isbtTermUsage"; }
        }

        public string GetCommandUrl(string strongNumber)
        {
            return string.Format("{0}:{1}", ProtocolName, strongNumber);
        }

        public bool IsProtocolCommand(string[] args)
        {
            return args.Length > 0 && args[0].StartsWith(ProtocolName, StringComparison.OrdinalIgnoreCase);
        }

        public void ExecuteCommand(string[] args)
        {
            if (args.Length == 0)
                throw new ArgumentNullException("args");

            if (!string.IsNullOrEmpty(SettingsManager.Instance.NotebookId_SupplementalBible))
            {
                string strongNumber = args[0].Split(new char[] { ':' }, StringSplitOptions.RemoveEmptyEntries)[1];
                Application oneNoteApp = new Application();
                string result;
                try
                {
                    oneNoteApp.FindPages(SettingsManager.Instance.NotebookId_SupplementalBible, strongNumber, out result, true, true, Consts.Constants.CurrentOneNoteSchema);
                }
                catch (COMException ex)
                {
                    if (ex.Message.EndsWith("0x80042019"))  // The query is invalid.
                    {
                        throw new Exception(BibleCommon.Resources.Constants.SearchQueryIsInvalid);
                    }
                }
                finally
                {
                    oneNoteApp = null;
                }
            }
        }        
    }
}
