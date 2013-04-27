using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using BibleCommon.Helpers;
using BibleCommon.Services;
using BibleCommon.Handlers;
using System.Threading;

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
            Application.ThreadException += new ThreadExceptionEventHandler(Application_ThreadException);

            var form = PrepareForRunning(args);

            if (form != null)
            {
                FormExtensions.RunSingleInstance(
                    BibleCommon.Consts.Constants.AnalyzeMutix, 
                    BibleCommon.Resources.Constants.MoreThanSingleInstanceRun, 
                    () => Application.Run(form));
            }

        }

        static void Application_ThreadException(object sender, ThreadExceptionEventArgs e)
        {
            FormLogger.LogError(e.Exception);
        }

        private static Form PrepareForRunning(params string[] args)
        {
            Form result = null;

            if (args.Contains(Consts.QuickAnalyze))
            {
                var handler = new QuickAnalyzeHandler();
                handler.ExecuteCommand();
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
