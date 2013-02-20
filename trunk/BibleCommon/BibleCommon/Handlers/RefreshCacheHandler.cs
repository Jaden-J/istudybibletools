using BibleCommon.Contracts;
using BibleCommon.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Office.Interop.OneNote;

namespace BibleCommon.Handlers
{
    public class RefreshCacheHandler : IProtocolHandler
    {
        public string ProtocolName
        {
            get { return "isbtRefreshCache:"; }
        }

        public string GetCommandUrl(string args)
        {
            return string.Format("{0}{1}", ProtocolName, "refreshCache");
        }

        public bool IsProtocolCommand(string[] args)
        {
            return args.Length > 0 && args[0].StartsWith(ProtocolName, StringComparison.OrdinalIgnoreCase);
        }

        public void ExecuteCommand(string[] args)
        {
            Application oneNoteApp = null;
            try
            {
                SettingsManager.Initialize();
                OneNoteProxy.Initialize();
            }
            catch (Exception ex)
            {
                FormLogger.LogError(ex);
            }
            finally
            {
                if (oneNoteApp != null)
                {
                    Marshal.ReleaseComObject(oneNoteApp);
                    oneNoteApp = null;
                }
            }
        }
    }
}
