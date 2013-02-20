using BibleCommon.Contracts;
using BibleCommon.Helpers;
using BibleCommon.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace BibleCommon.Handlers
{
    public class QuickAnalyzeHandler : IProtocolHandler
    {
        public string ProtocolName
        {
            get { return "isbtQuickAnalyze:"; }
        }

        public string GetCommandUrl(string args)
        {
            return string.Format("{0}{1}", ProtocolName, "currentPage");
        }

        public bool IsProtocolCommand(string[] args)
        {
            return args.Length > 0 && args[0].StartsWith(ProtocolName, StringComparison.OrdinalIgnoreCase);
        }

        public void ExecuteCommand(string[] args)
        {
            Microsoft.Office.Interop.OneNote.Application oneNoteApp = null;
            try
            {
                Logger.Init("QuickAnalyze");
                oneNoteApp = new Microsoft.Office.Interop.OneNote.Application();
                var currentPage = OneNoteUtils.GetCurrentPageInfo(ref oneNoteApp);
                using (NoteLinkManager noteLinkManager = new NoteLinkManager(oneNoteApp))
                {
                    noteLinkManager.LinkPageVerses(currentPage.NotebookId, currentPage.Id, NoteLinkManager.AnalyzeDepth.SetVersesLinks, false);
                    noteLinkManager.SetCursorOnNearestVerse(noteLinkManager.LastAnalyzedVerse);
                }
                OneNoteProxy.Instance.CommitAllModifiedPages(ref oneNoteApp, true, null, null, null);
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
