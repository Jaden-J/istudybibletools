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
            FormExtensions.RunSingleInstance(BibleCommon.Resources.Constants.MoreThanSingleInstanceRun, () =>
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);

                Form form = PrepareForRunning(args);

                if (form != null)
                {                    
                    Application.Run(form);
                }
            });
        }

        private static Form PrepareForRunning(params string[] args)
        {
            Form result = null;

            if (args.Contains(Consts.QuickAnalyze))
            {
                try
                {
                    var _oneNoteApp = new Microsoft.Office.Interop.OneNote.Application();
                    var currentPage = OneNoteUtils.GetCurrentPageInfo(_oneNoteApp);
                    using (NoteLinkManager noteLinkManager = new NoteLinkManager(_oneNoteApp))
                    {
                        noteLinkManager.LinkPageVerses(currentPage.SectionGroupId, currentPage.SectionId, currentPage.Id, NoteLinkManager.AnalyzeDepth.GetVersesLinks, false);
                    }
                    OneNoteProxy.Instance.CommitAllModifiedPages(_oneNoteApp, null, null, null);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                result = new MainForm();
            }

            return result;
        }

        public static void RunFromAnotherApp(params string[] args)
        {
             Form form = PrepareForRunning(args);

             if (form != null)
             {
                 form.ShowDialog();
                 form.Dispose();
             }
        }
    }
}
