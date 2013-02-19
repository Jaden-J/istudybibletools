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
                //Application.ThreadException += new ThreadExceptionEventHandler(Application_ThreadException);

                LanguageManager.SetThreadUICulture();

                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);

                bool silent;
                Action action;
                string mutexId;
                string moreThanSingleInstanceRunMessage;
                string additionalMutexId;
                Form form = PrepareForRunning(out silent, out moreThanSingleInstanceRunMessage, out action, out mutexId, out additionalMutexId, args);

                FormExtensions.RunSingleInstance(mutexId, moreThanSingleInstanceRunMessage, () =>
                {
                    try
                    {
                        if (action != null)
                            action();

                        if (form != null)                        
                            Application.Run(form);                                                
                    }
                    catch (Exception ex)
                    {
                        Logger.LogError(ex);
                    }
                }, silent, additionalMutexId);
            }
            catch (Exception ex)
            {
                FormLogger.LogError(ex);
            }
            finally
            {
                if (_oneNoteApp != null)
                {                    
                    _oneNoteApp = null;
                }
            }
        }

        static void Application_ThreadException(object sender, ThreadExceptionEventArgs e)
        {
            FormLogger.LogError(e.Exception);
        }

        private static Microsoft.Office.Interop.OneNote.Application _oneNoteApp;
        private static void CreateOneNoteAppIfNotExists()
        {
            if (_oneNoteApp == null)
                _oneNoteApp = new Microsoft.Office.Interop.OneNote.Application();
        }

        private static bool IsSystemConfigured()
        {
            CreateOneNoteAppIfNotExists();
            return SettingsManager.Instance.IsConfigured(ref _oneNoteApp);
        }

        private static Form PrepareForRunning(out bool silent, out string moreThanSingleInstanceRunMessage, out Action action, 
            out string mutexId, out string additionalMutexId, params string[] args)
        {
            action = null;
            moreThanSingleInstanceRunMessage = BibleCommon.Resources.Constants.MoreThanSingleInstanceRun;
            silent = false;
            Form result = null;
            mutexId = null;
            additionalMutexId = null;

            var rebuildDictionaryCacheHandler = new RebuildDictionaryFileCacheHandler();

            if (args.Contains(Consts.ShowModuleInfo) && IsSystemConfigured())
                result = new AboutModuleForm(SettingsManager.Instance.ModuleShortName, true);
            else if (args.Contains(Consts.ShowAboutProgram))
                result = new AboutProgramForm();
            else if (args.Contains(Consts.ShowSearchInDictionaries))
                result = new SearchInDictionariesForm();
            else if (args.Contains(Consts.ShowManual))
                OpenManual();
            else if (rebuildDictionaryCacheHandler.IsProtocolCommand(args))
            {
                rebuildDictionaryCacheHandler.ExecuteCommand(args);

                result = new MainForm(args);
                ((MainForm)result).ForceIndexDictionaryModuleName = rebuildDictionaryCacheHandler.ModuleShortName;
                ((MainForm)result).CommitChangesAfterLoad = true;
            }
            else if (args.Contains(Consts.RunOnOneNoteStarts))
            {
                silent = true;
                CreateOneNoteAppIfNotExists();
                if (SettingsManager.Instance.IsConfigured(ref _oneNoteApp))
                {
                    action = () =>
                        {
                            try
                            {
                                OneNoteLocker.LockBible(ref _oneNoteApp);
                                OneNoteLocker.LockSupplementalBible(ref _oneNoteApp);
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
                                        var mainForm = new MainForm(args);
                                        ((MainForm)mainForm).ToIndexBible = true;
                                        ((MainForm)mainForm).CommitChangesAfterLoad = true;
                                        Application.Run(mainForm);
                                    }
                                }
                            }
                        };

                    mutexId = BibleCommon.Consts.Constants.AnalyzeMutix;
                    additionalMutexId = BibleCommon.Consts.Constants.ParametersMutix;
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
                    CreateOneNoteAppIfNotExists();
                    OneNoteLocker.LockBible(ref _oneNoteApp);
                    OneNoteLocker.LockSupplementalBible(ref _oneNoteApp);
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
                    CreateOneNoteAppIfNotExists();
                    OneNoteLocker.UnlockBible(ref _oneNoteApp);
                    OneNoteLocker.UnlockSupplementalBible(ref _oneNoteApp);
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
                    CreateOneNoteAppIfNotExists();
                    OneNoteLocker.UnlockCurrentSection(ref _oneNoteApp);
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
                        using (var loadForm = new LoadForm())
                        {
                            loadForm.Show();
                            Application.DoEvents();

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

                            loadForm.Close();
                        }
                    }
                }
            }
            else
            {
                result = new MainForm(args);
            }

            if (result != null)
                mutexId = BibleCommon.Consts.Constants.ParametersMutix;

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
                Action action;
                string mutexId;
                string additionalMutexId;
                Form form = PrepareForRunning(out silent, out moreThanSingleInstanceRunMessage, out action, out mutexId, out additionalMutexId, args);

                if (action != null)
                    action();

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
                {                    
                    _oneNoteApp = null;
                }
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
