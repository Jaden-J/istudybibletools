using BibleCommon.Contracts;
using BibleCommon.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Office.Interop.OneNote;
using BibleCommon.Common;
using BibleCommon.Helpers;

namespace BibleCommon.Handlers
{
    public class RefreshCacheHandler : IProtocolHandler
    {
        public enum RefreshCacheMode
        {
            RefreshApplicationCache,
            RefreshAnalyzedVersesCache
        }

        private const string _protocolName = "isbtRefreshCache:";

        public string ProtocolName
        {
            get { return _protocolName; }
        }

        public static string GetCommandUrlStatic(RefreshCacheMode mode)
        {
            return string.Format("{0}{1}", _protocolName, mode);
        }

        public string GetCommandUrl(string args)
        {
            var mode = !string.IsNullOrEmpty(args) ? (RefreshCacheMode)Enum.Parse(typeof(RefreshCacheMode), args) : RefreshCacheMode.RefreshApplicationCache;
            return GetCommandUrlStatic(mode);
        }

        public bool IsProtocolCommand(params string[] args)
        {
            return args.Length > 0 && args[0].StartsWith(ProtocolName, StringComparison.OrdinalIgnoreCase);
        }

        public RefreshCacheMode CacheMode { get; set; }

        public void ExecuteCommand(params string[] args)
        {
            Application oneNoteApp = null;
            try
            {
                CacheMode = args.Length > 1 ? (RefreshCacheMode)Enum.Parse(typeof(RefreshCacheMode), args[1]) : RefreshCacheMode.RefreshApplicationCache;

                switch (CacheMode)
                {
                    case RefreshCacheMode.RefreshApplicationCache:
                        oneNoteApp = OneNoteUtils.CreateOneNoteAppSafe();  // для разгона
                        SettingsManager.Initialize();
                        ApplicationCache.Initialize();
                        break;
                    case RefreshCacheMode.RefreshAnalyzedVersesCache:
                        // дальнейшая обработка осуществляется в CommandForm.ProcessCommandLine()
                        break;
                }                

                //BibleCommon.Resources.Constants.Culture = LanguageManager.UserLanguage;
            }
            catch (NotConfiguredException)
            { }
            catch (ModuleIsUndefinedException)
            { }
            catch (Exception ex)
            {
                FormLogger.LogError(ex);
            }
            finally
            {
                OneNoteUtils.ReleaseOneNoteApp(ref oneNoteApp);
            }
        }  
    }
}
