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
        private const int SleepMilisecondsOnOneNoteStarts = 3000;

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
                Application.ThreadException += new ThreadExceptionEventHandler(Application_ThreadException);

                bool silent;
                Action action;
                string mutexId;
                string moreThanSingleInstanceRunMessage;
                string additionalMutexId;
                var form = PrepareForRunning(out silent, out moreThanSingleInstanceRunMessage, out action, out mutexId, out additionalMutexId, args);

                FormExtensions.RunSingleInstance(mutexId, moreThanSingleInstanceRunMessage, () =>
                {
                    try
                    {
                        if (action != null)
                            action();

                        if (form != null)
                        {   
                            Application.Run(form);
                        }
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
                OneNoteUtils.ReleaseOneNoteApp(ref _oneNoteApp);
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
                _oneNoteApp = OneNoteUtils.CreateOneNoteAppSafe();
        }

        private static bool IsSystemConfigured()
        {
            CreateOneNoteAppIfNotExists();
            return SettingsManager.Instance.IsConfigured(ref _oneNoteApp);
        }

        private static Form PrepareForRunning(out bool silent, out string moreThanSingleInstanceRunMessage, out Action action, 
            out string mutexId, out string additionalMutexId, params string[] args)
        {            
            moreThanSingleInstanceRunMessage = BibleCommon.Resources.Constants.MoreThanSingleInstanceRun;
            action = null;
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
                result = RebuildDictionaryCache(args, rebuildDictionaryCacheHandler);
            else if (args.Contains(Consts.RunOnOneNoteStarts))
                result = OnOneNoteStarts(args, out action, out mutexId, out additionalMutexId, out silent);                   
            else if (args.Contains(Consts.LockAllBible))
                LockAllBible();               
            else if (args.Contains(Consts.UnlockAllBible))
                UnlockAllBible();               
            else if (args.Contains(Consts.UnlockBibleSection))
                UnlockBibleSection();               
            else if (args.Length == 1)
                result = TryToLoadModule(args, ref moreThanSingleInstanceRunMessage);               
            else
                result = new MainForm(args);

            if (result != null)
                mutexId = BibleCommon.Consts.Constants.ParametersMutix;

            return result;
        }

        private static Form RebuildDictionaryCache(string[] args, RebuildDictionaryFileCacheHandler rebuildDictionaryCacheHandler)
        {
            rebuildDictionaryCacheHandler.ExecuteCommand(args);

            var result = new MainForm(args);
            ((MainForm)result).ForceIndexDictionaryModuleName = rebuildDictionaryCacheHandler.ModuleShortName;
            ((MainForm)result).CommitChangesAfterLoad = true;
            ((MainForm)result).NotAskToIndexBible = true;   // а то выглядит непонятно, когда нас попросили перестроить кэш словаря и тут же сразу просят проиндексировать Библию

            return result;
        }

        private static Form TryToLoadModule(string[] args, ref string moreThanSingleInstanceRunMessage)
        {
            var result = new MainForm(args);

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

            return result;
        }

        private static void UnlockBibleSection()
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

        private static void LockAllBible()
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

        private static void UnlockAllBible()
        {
            try
            {
                CreateOneNoteAppIfNotExists();
                OneNoteLocker.UnlockBible(ref _oneNoteApp);
            }
            catch (NotSupportedException)
            {
                ShowMessage(BibleCommon.Resources.Constants.SkyDriveBibleIsNotSupportedForLock);
            }

            try
            {
                OneNoteLocker.UnlockSupplementalBible(ref _oneNoteApp);
            }
            catch (NotSupportedException ex)
            {
                Logger.LogError(ex);
            }
        }

        private static Form OnOneNoteStarts(string[] args, out Action action, out string mutexId, out string additionalMutexId, out bool silent)
        {
            action = null;
            Form result = null;
            mutexId = null;
            additionalMutexId = null;
            silent = true;

            Thread.Sleep(SleepMilisecondsOnOneNoteStarts);      // немного времени ждём, пока OneNote загрузится
            
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

                    SettingsManager.Instance.TryToUpdateVersion();      // если надо - применяем обновления и обновляем версию 

                    var vsm = new VersionOnServerManager();
                    if (vsm.NeedToUpdate() && !ShownMessagesManager.GetMessageWasShown(ShownMessagesManager.MessagesCodes.NewVersionInfo))
                    {
                        var htmlForm = new HtmlMessageForm(ShownMessagesManager.MessagesCodes.NewVersionInfo, vsm.ReleaseInfo, BibleCommon.Resources.Constants.NewVersionAvailable);
                        htmlForm.ShowDialog();
                    }

                    if (!BibleVersesLinksCacheManager.CacheIsActive(SettingsManager.Instance.NotebookId_Bible))
                    {
                        var minutes = MainForm.GetMinutesForBibleVersesCacheGenerating();
                        using (var form = new MessageForm(string.Format(BibleCommon.Resources.Constants.IndexBibleQuestionAtStartUp, minutes), BibleCommon.Resources.Constants.Warning,
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

                mutexId = BibleCommon.Consts.Constants.AnalysisMutix;
                additionalMutexId = BibleCommon.Consts.Constants.ParametersMutix;

                if (_oneNoteApp.Windows.CurrentWindow != null)        
                    Process.Start(RefreshCacheHandler.GetCommandUrlStatic(RefreshCacheHandler.RefreshCacheMode.RefreshApplicationCache));  // инициализируем кэш
            }
            else
            {
                result = new MainForm(args);
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
                OneNoteUtils.ReleaseOneNoteApp(ref _oneNoteApp);
            }
        }        

        private static bool OpenManual()
        {
            var path = Path.GetDirectoryName(Path.GetDirectoryName(Utils.GetCurrentDirectory()));

            var files = Directory.GetFiles(path, string.Format("Instruction*{0}*", LanguageManager.GetCurrentCultureInfo().LCID));
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
