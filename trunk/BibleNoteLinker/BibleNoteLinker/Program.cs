using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using BibleCommon.Helpers;
using BibleCommon.Services;

namespace BibleNoteLinker
{
    public static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(params string[] args)
        {

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            var form = PrepareForRunning(args);

            if (form != null)
            {
                FormExtensions.RunSingleInstance(
                    BibleCommon.Consts.Constants.AnalyzeMutix, 
                    BibleCommon.Resources.Constants.MoreThanSingleInstanceRun, 
                    () => Application.Run(form));
            }

        }

        private static Form PrepareForRunning(params string[] args)
        {
            Form result = null;

            if (args.Contains(Consts.QuickAnalyze))
            {
                try
                {
                    Logger.Init("QuickAnalyze");
                    var _oneNoteApp = new Microsoft.Office.Interop.OneNote.Application();
                    var currentPage = OneNoteUtils.GetCurrentPageInfo(ref _oneNoteApp);
                    using (NoteLinkManager noteLinkManager = new NoteLinkManager(_oneNoteApp))
                    {
                        noteLinkManager.LinkPageVerses(currentPage.NotebookId, currentPage.Id, NoteLinkManager.AnalyzeDepth.SetVersesLinks, false);
                        noteLinkManager.SetCursorOnNearestVerse(noteLinkManager.LastAnalyzedVerse);                        
                    }
                    OneNoteProxy.Instance.CommitAllModifiedPages(ref _oneNoteApp, true, null, null, null);
                }
                catch (Exception ex)
                {
                    FormLogger.LogError(ex);
                }
            }
            else
            {
                result = new MainForm();
            }

            return result;
        }

        private static bool _firstLoad = true;
        public static void RunFromAnotherApp(params string[] args)
        {
            try
            {
                if (_firstLoad)
                {
                    try
                    {
                        Application.EnableVisualStyles();
                        Application.SetCompatibleTextRenderingDefault(false);
                    }
                    catch { }
                    _firstLoad = false;
                }

                Form form = PrepareForRunning(args);

                if (form != null)
                {
                    form.ShowDialog();
                    form.Dispose();
                }
            }
            catch (Exception ex)
            {
                FormLogger.LogError(ex);
            }
        }
    }
}
