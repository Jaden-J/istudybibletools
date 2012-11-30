using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.OneNote;
using BibleCommon.Services;
using BibleCommon.Consts;
using BibleCommon.Common;
using BibleCommon.Helpers;
using System.IO;
using System.Reflection;


namespace BibleConfigurator.Tools
{    
    // он не IDisposable, потому что рабоатет асинхронно. И сам по окончанию диспозит себя. 
    public class BackupManager
    {
        public const string OneNotePackageExtension = ".onepkg";
        private Application _oneNoteApp;
        private MainForm _form;
        private FileSystemWatcher _fileWatcher;
        private int _processedNotebooksCount = 0;
        private int _notebooksCount = 0;
        private string _tempFolderPath;
        private string _targetFilePath;
        private volatile List<String> _notebookNames = new List<string>();
        private DateTime _startTime;
        private static object _locker = new object();

        public BackupManager(Application oneNoteApp, MainForm form)
        {
            _oneNoteApp = oneNoteApp;
            _form = form;
        }

        public void Backup(string filePath)
        {
            if (!SettingsManager.Instance.IsConfigured(_oneNoteApp))
            {
                FormLogger.LogError(BibleCommon.Resources.Constants.Error_SystemIsNotConfigured);
                return;
            }   

            try
            {
                BibleCommon.Services.Logger.Init("BackupManager");

                _startTime = DateTime.Now;
                BibleCommon.Services.Logger.LogMessage("{0}: {1}",BibleCommon.Resources.Constants.StartTime,  _startTime.ToLongTimeString());                

                var notebookIds = GetDistinctNotebooksIds();
                _notebooksCount = notebookIds.Count();

                string initMessage = BibleCommon.Resources.Constants.BackupStartInfo;
                _form.PrepareForLongProcessing(_notebooksCount + 2, 1, initMessage);
                _form.PerformProgressStep(initMessage);
                BibleCommon.Services.Logger.LogMessage(initMessage);
                System.Windows.Forms.Application.DoEvents();

                _tempFolderPath = Path.Combine(Utils.GetTempFolderPath(), "BackUp");

                if (!Directory.Exists(_tempFolderPath))
                    Directory.CreateDirectory(_tempFolderPath);

                _targetFilePath = filePath;

                CleanTempFolder();

                InitializeFileWatcher();

                foreach (string id in notebookIds)
                {
                    string notebookName = OneNoteUtils.GetHierarchyElementNickname(_oneNoteApp, id) + OneNotePackageExtension;

                    _notebookNames.Add(notebookName);

                    BackupNotebook(id, notebookName);                    
                    
                    if (_form.StopLongProcess)
                        throw new ProcessAbortedByUserException();
                }                
            }
            catch (ProcessAbortedByUserException)
            {
                BibleCommon.Services.Logger.LogMessage("Process aborted by user");
                CloseResources(false);
            }            
        }

        private void CloseResources(bool successefully)
        {
            _fileWatcher.EnableRaisingEvents = false;
            _fileWatcher.Dispose();           

            try
            {
                CleanTempFolder();
            }
            catch (Exception ex)
            {
                BibleCommon.Services.Logger.LogError(ex.Message);
            }

            BibleCommon.Services.Logger.LogMessage(" {0}", DateTime.Now.Subtract(_startTime));
            BibleCommon.Services.Logger.Done();

            if (successefully)
            {
                string finalMessage = BibleCommon.Resources.Constants.BackupManagerFinishMessage;
                _form.LongProcessingDone(finalMessage);
                BibleCommon.Services.Logger.LogMessage(finalMessage);
            }

            _oneNoteApp = null;
            _form = null;
        }

        private void InitializeFileWatcher()
        {
            _fileWatcher = new FileSystemWatcher(_tempFolderPath, "*" + OneNotePackageExtension);

            _fileWatcher.Changed += new FileSystemEventHandler(_fileWatcher_Changed);            
            _fileWatcher.EnableRaisingEvents = true;
        }

        void _fileWatcher_Changed(object sender, FileSystemEventArgs e)
        {
            try
            {
                if (_form.StopLongProcess)
                    throw new ProcessAbortedByUserException();

                if (new FileInfo(e.FullPath).Length > 0)
                {
                    if (_notebookNames.Contains(e.Name))
                    {
                        lock (_locker)
                        {
                            if (_notebookNames.Contains(e.Name))
                            {
                                _notebookNames.Remove(e.Name);
                                string message = string.Format(BibleCommon.Resources.Constants.BackupManagerNotebookCompleted,
                                    Path.GetFileNameWithoutExtension(e.Name));
                                _form.PerformProgressStep(message);
                                BibleCommon.Services.Logger.LogMessage(message);

                                if (++_processedNotebooksCount >= _notebooksCount)
                                {
                                    PackfilesToZip();
                                    CloseResources(true);
                                }
                            }
                        }
                    }
                }
            }
            catch (ProcessAbortedByUserException)
            {
                BibleCommon.Services.Logger.LogMessage("Process aborted by user");

                CloseResources(false);
            }
        }        

        private void BackupNotebook(string notebookId, string notebookName)
        {
            _oneNoteApp.Publish(notebookId, Path.Combine(_tempFolderPath,
                notebookName), PublishFormat.pfOneNotePackage);
        }

        private IEnumerable<string> GetDistinctNotebooksIds()
        {
            return new List<string>(SettingsManager.Instance.SelectedNotebooksForAnalyze) 
            {
                SettingsManager.Instance.NotebookId_Bible,
                //SettingsManager.Instance.NotebookId_BibleStudy,
                //SettingsManager.Instance.NotebookId_BibleComments //,
                //SettingsManager.Instance.NotebookId_BibleNotesPages                
            }.Distinct();
        }

        private void PackfilesToZip()
        {
            string message = BibleCommon.Resources.Constants.BackupManagerToZipArchive;
            _form.PerformProgressStep(message);
            BibleCommon.Services.Logger.LogMessage(message);

            try
            {
                ZipLibHelper.PackfilesToZip(_tempFolderPath, _targetFilePath);
            }
            catch (Exception ex)
            {
                FormLogger.LogError(ex);                
            }
        }

        private void CleanTempFolder()
        {
            foreach (string file in Directory.GetFiles(_tempFolderPath))
            {
                File.Delete(file);
            }
        }       
    }    
}
