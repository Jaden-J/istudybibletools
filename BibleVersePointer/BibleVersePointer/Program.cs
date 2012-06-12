using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using BibleCommon.Helpers;

namespace BibleVersePointer
{
    static class Program
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
            Form result = new MainForm();

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
