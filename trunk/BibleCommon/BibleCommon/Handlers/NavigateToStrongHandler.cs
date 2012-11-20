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
            try
            {
                TryExecuteCommand(args);
            }
            catch (Exception ex)
            {
                FormLogger.LogError(ex);
            }           
        }

        private void TryExecuteCommand(string[] args)
        {
            if (args.Length == 0)
                throw new ArgumentNullException("args");

            var verseArgs = Uri.UnescapeDataString(args[0]
                                .Split(new char[] { ':' })[1])
                                .Split(new char[] { ';' });

            var strongNumber = verseArgs[0];
            var moduleShortName = verseArgs[1];

            Application oneNoteApp = new Application();

            try
            {
                var strongTermLink = OneNoteProxy.Instance.GetDictionaryTermLink(strongNumber, moduleShortName);
                
                try
                {
                    oneNoteApp.NavigateTo(strongTermLink.PageId, strongTermLink.ObjectId);
                }
                finally
                {
                    oneNoteApp = null;
                }
            }
            catch (Exception ex)  // todo
            {
                if (!DictionaryTermsCacheManager.CacheIsActive(moduleShortName))
                {
                    if (string.IsNullOrEmpty(SettingsManager.Instance.GetValidDictionariesNotebookId(oneNoteApp, true)))
                        throw new Exception(BibleCommon.Resources.Constants.DictionariesNotebookNotFound);

                    throw new Exception(BibleCommon.Resources.Constants.DictionaryCacheFileNotFound); 
                }

                throw;
            }
        }

        string IProtocolHandler.GetCommandUrl(string args)
        {
            return string.Format("{0}:{1}", ProtocolName, args);
        }      
    }
}
