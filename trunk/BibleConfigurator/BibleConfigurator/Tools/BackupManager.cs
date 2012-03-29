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
using ICSharpCode.SharpZipLib.Zip;

namespace BibleConfigurator.Tools
{    
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
                Logger.LogError(Constants.Error_SystemIsNotConfigures);
                return;
            }   

            try
            {
                BibleCommon.Services.Logger.Init("BackupManager");

                IEnumerable<string> notebookIds = GetDistinctNotebooksIds();
                _notebooksCount = notebookIds.Count();

                _form.PrepareForExternalProcessing(_notebooksCount + 1, 1, "Старт создания резервной копии данных");

                _tempFolderPath = GetTempFolderPath();

                _targetFilePath = filePath;

                CleanTempFolder();

                InitializeFileWatcher();

                foreach (string id in notebookIds)
                {
                    string notebookName = OneNoteUtils.GetHierarchyElementName(_oneNoteApp, id) + OneNotePackageExtension;

                    _notebookNames.Add(notebookName);

                    BackupNotebook(id, notebookName);                    
                    
                    if (_form.StopExternalProcess)
                        throw new ProcessAbortedByUserException();
                }                
            }
            catch (ProcessAbortedByUserException)
            {
                BibleCommon.Services.Logger.LogMessage("Process aborted by user");
                Finalize(false);
            }
            finally
            {
                BibleCommon.Services.Logger.Done();
            }
        }

        private void Finalize(bool successefully)
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

            BibleCommon.Services.Logger.Done();

            if (successefully)
                _form.ExternalProcessingDone("Создание резервной копии данных успешно завершено.");
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
                if (_form.StopExternalProcess)
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
                                _form.PerformProgressStep(string.Format("Резервная копия записной книжки '{0}' успешно создана.", 
                                    Path.GetFileNameWithoutExtension(e.Name)));

                                if (++_processedNotebooksCount >= _notebooksCount)
                                {
                                    PackfilesToZip();
                                    Finalize(true);
                                }
                            }
                        }
                    }
                }
            }
            catch (ProcessAbortedByUserException)
            {
                BibleCommon.Services.Logger.LogMessage("Process aborted by user");

                Finalize(false);
            }
        }        

        private void BackupNotebook(string notebookId, string notebookName)
        {
            _oneNoteApp.Publish(notebookId, Path.Combine(_tempFolderPath,
                notebookName), PublishFormat.pfOneNotePackage);
        }

        private IEnumerable<string> GetDistinctNotebooksIds()
        {
            return new List<string>() 
            {
                SettingsManager.Instance.NotebookId_Bible,
                SettingsManager.Instance.NotebookId_BibleComments,
                SettingsManager.Instance.NotebookId_BibleNotesPages,
                SettingsManager.Instance.NotebookId_BibleStudy
            }.Distinct();
        }

        private void PackfilesToZip()
        {
            _form.PerformProgressStep("Упаковывание файлов в .zip архив");
            try
            {
                // Depending on the directory this could be very large and would require more attention
                // in a commercial package.
                string[] filenames = Directory.GetFiles(_tempFolderPath);

                // 'using' statements guarantee the stream is closed properly which is a big source
                // of problems otherwise.  Its exception safe as well which is great.
                using (ZipOutputStream s = new ZipOutputStream(File.Create(_targetFilePath)))
                {
                    s.SetLevel(9); // 0 - store only to 9 - means best compression

                    byte[] buffer = new byte[4096];

                    foreach (string file in filenames)
                    {
                        // Using GetFileName makes the result compatible with XP
                        // as the resulting path is not absolute.
                        ZipEntry entry = new ZipEntry(Path.GetFileName(file));

                        // Setup the entry data as required.

                        // Crc and size are handled by the library for seakable streams
                        // so no need to do them here.

                        // Could also use the last write time or similar for the file.
                        entry.DateTime = DateTime.Now;
                        s.PutNextEntry(entry);

                        using (FileStream fs = File.OpenRead(file))
                        {

                            // Using a fixed size buffer here makes no noticeable difference for output
                            // but keeps a lid on memory usage.
                            int sourceBytes;
                            do
                            {
                                sourceBytes = fs.Read(buffer, 0, buffer.Length);
                                s.Write(buffer, 0, sourceBytes);
                            } while (sourceBytes > 0);
                        }
                    }

                    // Finish/Close arent needed strictly as the using statement does this automatically

                    // Finish is important to ensure trailing information for a Zip file is appended.  Without this
                    // the created file would be invalid.
                    s.Finish();

                    // Close is important to wrap things up and unlock the file.
                    s.Close();
                }
            }
            catch (Exception ex)
            {
                Logger.LogError(ex.Message);                

                // No need to rethrow the exception as for our purposes its handled.
            }
        }

        private void CleanTempFolder()
        {
            foreach (string file in Directory.GetFiles(_tempFolderPath))
            {
                File.Delete(file);
            }
        }

        private static string GetTempFolderPath()
        {
            string s = Path.Combine(Path.GetDirectoryName(Path.GetDirectoryName(Utils.GetCurrentDirectory())), Consts.TempDirectory);
            if (!Directory.Exists(s))
                Directory.CreateDirectory(s);

            return s;
        }
    }    
}
