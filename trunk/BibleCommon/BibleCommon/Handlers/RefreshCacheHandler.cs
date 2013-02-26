using BibleCommon.Contracts;
using BibleCommon.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Office.Interop.OneNote;
using BibleCommon.Common;

namespace BibleCommon.Handlers
{
    public class RefreshCacheHandler : IProtocolHandler
    {
        private const string _protocolName = "isbtRefreshCache:";

        public string ProtocolName
        {
            get { return _protocolName; }
        }

        public static string GetCommandUrlStatic()
        {
            return string.Format("{0}{1}", _protocolName, "refreshCache");
        }

        public string GetCommandUrl(string args)
        {
            return GetCommandUrlStatic();
        }

        public bool IsProtocolCommand(params string[] args)
        {
            return args.Length > 0 && args[0].StartsWith(ProtocolName, StringComparison.OrdinalIgnoreCase);
        }

        public void ExecuteCommand(params string[] args)
        {
            Application oneNoteApp = null;
            try
            {
                oneNoteApp = new Application();  // для разгона

                SettingsManager.Initialize();
                OneNoteProxy.Initialize();

                //BibleCommon.Resources.Constants.Culture = LanguageManager.UserLanguage;
            }
            catch (NotConfiguredException)
            { }
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
