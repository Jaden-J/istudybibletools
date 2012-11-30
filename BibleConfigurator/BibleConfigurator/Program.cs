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
using BibleCommon.Handlers;
using BibleCommon.UI.Forms;


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

                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);

                bool silent;                
                string moreThanSingleInstanceRunMessage;
                Form form = PrepareForRunning(out silent, out moreThanSingleInstanceRunMessage, args);

                if (form != null)
                {
                    FormExtensions.RunSingleInstance(form, moreThanSingleInstanceRunMessage, () =>
                    {
                        try
                        {
                            Application.Run(form);
                        }
                        catch (Exception ex)
                        {
                            Logger.LogError(ex);
                        }
                    }, silent || args.Contains(Consts.RunOnOneNoteStarts));
                }

            }
            catch (Exception ex)
            {
                FormLogger.LogError(ex);
            }
            finally
            {
                if (_oneNoteApp != null)
                    _oneNoteApp = null;
            }
        }

        private static Form PrepareForRunning(out bool silent, out string moreThanSingleInstanceRunMessage, params string[] args)
        {
            moreThanSingleInstanceRunMessage = BibleCommon.Resources.Constants.MoreThanSingleInstanceRun;
            silent = false;
            Form result = null;

            var strongProtocolHandler = new FindVersesWithStrongNumberHandler();
            var navToStrongHandler = new NavigateToStrongHandler();

            if (args.Contains(Consts.ShowModuleInfo) && SettingsManager.Instance.IsConfigured(OneNoteApp))
                result = new AboutModuleForm(SettingsManager.Instance.ModuleShortName, true);
            else if (args.Contains(Consts.ShowAboutProgram))
                result = new AboutProgramForm();
            else if (args.Contains(Consts.ShowSearchInDictionaries))
                result = new SearchInDictionariesForm();
            else if (args.Contains(Consts.ShowManual))
                OpenManual();
            else if (strongProtocolHandler.IsProtocolCommand(args))
                strongProtocolHandler.ExecuteCommand(args);
            else if (navToStrongHandler.IsProtocolCommand(args))
            {
                if (!navToStrongHandler.ExecuteCommand(args))
                {
                    result = new MainForm(args);
                    ((MainForm)result).ForceIndexDictionaryModuleName = navToStrongHandler.ModuleShortName;
                    ((MainForm)result).CommitChangesAfterLoad = true;
                }
            }            
            else if (args.Contains(Consts.RunOnOneNoteStarts))
            {                
                if (SettingsManager.Instance.IsConfigured(OneNoteApp))
                {
                    try
                    {
                        OneNoteLocker.LockBible(OneNoteApp);
                        OneNoteLocker.LockSupplementalBible(OneNoteApp);
                    }
                    catch (NotSupportedException)
                    {
                        Logger.LogError("Locking is not supported for this notebook");
                    }

                    if (!BibleVersesLinksCacheManager.CacheIsActive(SettingsManager.Instance.NotebookId_Bible))
                    {
                        using (var form = new MessageForm(BibleCommon.Resources.Constants.IndexBibleQuestionAtStartUp, BibleCommon.Resources.Constants.Warning,
                            MessageBoxButtons.YesNo, MessageBoxIcon.Question))
                        {
                            if (form.ShowDialog() == DialogResult.Yes)
                            {
                                result = new MainForm(args);
                                ((MainForm)result).ToIndexBible = true;
                                ((MainForm)result).CommitChangesAfterLoad = true;
                            }
                        }
                    }
                }
                else
                {
                    result = new MainForm(args);
                }
            }
            else if (args.Contains(Consts.LockAllBible))
            {
                try
                {
                    OneNoteLocker.LockBible(OneNoteApp);
                    OneNoteLocker.LockSupplementalBible(OneNoteApp);
                }
                catch (NotSupportedException)
                {
                    ShowMessage(BibleCommon.Resources.Constants.SkyDriveBibleIsNotSupportedForLock);
                }
            }
            else if (args.Contains(Consts.UnlockAllBible))
            {
                try
                {
                    OneNoteLocker.UnlockBible(OneNoteApp);
                    OneNoteLocker.UnlockSupplementalBible(OneNoteApp);
                }
                catch (NotSupportedException)
                {
                    ShowMessage(BibleCommon.Resources.Constants.SkyDriveBibleIsNotSupportedForLock);
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
                    ShowMessage(BibleCommon.Resources.Constants.SkyDriveBibleIsNotSupportedForLock);
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
                            moreThanSingleInstanceRunMessage = BibleCommon.Resources.Constants.ReopenParametersToSeeChanges;
                            //silent = true;
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

            return result;
        }

        private static bool _firstLoad = true;
        //public static void RunFromAnotherAppNotDialog(params string[] args)
        //{
        //    if (_firstLoad)
        //    {
        //        Application.EnableVisualStyles();
        //        Application.SetCompatibleTextRenderingDefault(false);
        //        _firstLoad = false;
        //    }

        //    Form form = PrepareForRunning(args);            

        //    if (form != null)
        //    {
        //        Application.Run(form);
        //    }
        //}

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

                bool silent;
                string moreThanSingleInstanceRunMessage;
                Form form = PrepareForRunning(out silent, out moreThanSingleInstanceRunMessage, args);

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
            finally
            {
                if (_oneNoteApp != null)
                    _oneNoteApp = null;
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

        private static void ShowMessage(string message)
        {
            FormLogger.LogMessage(message);
        }
    }
}
