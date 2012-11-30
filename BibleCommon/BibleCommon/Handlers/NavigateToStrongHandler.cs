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

        /// <summary>
        /// Доступно только после вызова ExecuteCommand()
        /// </summary>
        public string ModuleShortName { get; set; }

        /// <summary>
        /// Доступно только после вызова ExecuteCommand()
        /// </summary>
        public string StrongNumber { get; set; }

        public static string GetCommandUrlStatic(string strongNumber, string moduleShortName)
        {
            return string.Format("{0}:{1};{2}", _protocolName, strongNumber, moduleShortName);
        }      

        public bool IsProtocolCommand(string[] args)
        {
            return args.Length > 0 && args[0].StartsWith(ProtocolName, StringComparison.OrdinalIgnoreCase);
        }

        public bool ExecuteCommand(string[] args)
        {
            try
            {
                return TryExecuteCommand(args);
            }
            catch (Exception ex)
            {
                FormLogger.LogError(ex);
                return true;
            }           
        }

        private bool TryExecuteCommand(string[] args)
        {
            if (args.Length == 0)
                throw new ArgumentNullException("args");

            var verseArgs = Uri.UnescapeDataString(args[0]
                                .Split(new char[] { ':' })[1])
                                .Split(new char[] { ';' });

            StrongNumber = verseArgs[0];
            ModuleShortName = verseArgs[1];

            Application oneNoteApp = new Application();

            try
            {
                var strongTermLink = OneNoteProxy.Instance.GetDictionaryTermLink(StrongNumber, ModuleShortName);

                return DictionaryManager.GoToTerm(oneNoteApp, strongTermLink);
            }
            catch (Exception)
            {
                if (!DictionaryTermsCacheManager.CacheIsActive(ModuleShortName))
                {
                    if (string.IsNullOrEmpty(SettingsManager.Instance.GetValidDictionariesNotebookId(oneNoteApp, true)))
                        throw new Exception(BibleCommon.Resources.Constants.DictionariesNotebookNotFound);

                    throw new Exception(BibleCommon.Resources.Constants.DictionaryCacheFileNotFound);
                }

                throw;
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
