using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using BibleCommon.Common;
using BibleConfigurator.ModuleConverter;
using System.Text;
using BibleCommon.Services;
using BibleCommon.Helpers;
using System.IO;
using System.Diagnostics;
using System.Xml.XPath;
using System.Threading;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Xml;
using System.Xml.Linq;
using BibleCommon.Consts;


namespace BibleConfigurator
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(params string[] args)
        {
            try
            {
                LanguageManager.SetThreadUICulture();

                string message = BibleCommon.Resources.Constants.MoreThanSingleInstanceRun;
                if (args.Length == 1 && File.Exists(args[0]))
                    message += " " + BibleCommon.Resources.Constants.LoadMofuleInExistingInstance;

                FormExtensions.RunSingleInstance(message, () =>
                {
                    Application.EnableVisualStyles();
                    Application.SetCompatibleTextRenderingDefault(false);

                    Form form = PrepareForRunning(args);

                    if (form != null)
                    {                        
                        Application.Run(form);
                    }
                }, args.Contains(Consts.RunOnOneNoteStarts));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }             

        private static Form PrepareForRunning(params string[] args)
        {
            Form result = null;

            try
            {
                if (args.Contains(Consts.ShowModuleInfo) && SettingsManager.Instance.IsConfigured(OneNoteApp))
                    result = new AboutModuleForm(SettingsManager.Instance.ModuleName, true);
                else if (args.Contains(Consts.ShowAboutProgram))
                    result = new AboutProgramForm();                
                else if (args.Contains(Consts.ShowManual))
                {
                    OpenManual();
                }
                else if (args.Contains(Consts.RunOnOneNoteStarts))
                {
                    if (SettingsManager.Instance.IsConfigured(OneNoteApp))
                    {
                        try
                        {
                            OneNoteLocker.LockAllBible(OneNoteApp);
                        }
                        catch (NotSupportedException)
                        {
                            //todo: log it
                        }
                    }
                    else
                        result = new MainForm(args);
                }
                else if (args.Contains(Consts.LockAllBible))
                {
                    try
                    {
                        OneNoteLocker.LockAllBible(OneNoteApp);
                    }
                    catch (NotSupportedException)
                    {
                        MessageBox.Show(BibleCommon.Resources.Constants.SkyDriveBibleIsNotSupportedForLock);
                    }
                }
                else if (args.Contains(Consts.UnlockAllBible))
                {
                    try
                    {
                        OneNoteLocker.UnlockAllBible(OneNoteApp);
                    }
                    catch (NotSupportedException)
                    {
                        MessageBox.Show(BibleCommon.Resources.Constants.SkyDriveBibleIsNotSupportedForLock);
                    }
                }
                else if (args.Contains(Consts.UnlockBibleSection))
                {
                    try
                    {
                        OneNoteLocker.UnlockCurrentSection(OneNoteApp);
                    }
                    catch (NotSupportedException)
                    {
                        MessageBox.Show(BibleCommon.Resources.Constants.SkyDriveBibleIsNotSupportedForLock);
                    }
                }
                else if (args.Length == 1)
                {
                    result = new MainForm(args);

                    if (!string.IsNullOrEmpty(args[0]))
                    {
                        string moduleFilePath = args[0];
                        if (File.Exists(moduleFilePath))
                        {
                            bool moduleWasAdded;
                            bool needToReload = ((MainForm)result).AddNewModule(moduleFilePath, out moduleWasAdded);
                            if (moduleWasAdded)
                            {
                                ((MainForm)result).ShowModulesTabAtStartUp = true;
                                ((MainForm)result).NeedToSaveChangesAfterLoadingModuleAtStartUp = needToReload;
                            }
                            else
                                result = null;
                        }
                    }
                }
                else
                {
                    result = new MainForm(args);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            if (_oneNoteApp != null)
                _oneNoteApp = null;

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

        private static Microsoft.Office.Interop.OneNote.Application _oneNoteApp;
        private static Microsoft.Office.Interop.OneNote.Application OneNoteApp
        {
            get
            {
                if (_oneNoteApp == null)
                    _oneNoteApp = new Microsoft.Office.Interop.OneNote.Application();

                return _oneNoteApp;
            }
        }

  


        private static bool OpenManual()
        {
            var path = Path.GetDirectoryName(Path.GetDirectoryName(Utils.GetCurrentDirectory()));

            var files = Directory.GetFiles(path, string.Format("Instruction*{0}*", LanguageManager.UserLanguage.LCID));
            if (files.Length == 0)
                files = Directory.GetFiles(path, string.Format("Instruction*{0}*", LanguageManager.DefaultLCID));

            if (files.Length == 1)
            {
                Process.Start(files[0]);
                return true;
            }

            return false;
        }


       



  
    }
}
