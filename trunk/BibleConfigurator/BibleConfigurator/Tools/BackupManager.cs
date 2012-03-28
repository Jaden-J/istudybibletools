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

namespace BibleConfigurator.Tools
{    
    public class BackupManager
    {
        private Application _oneNoteApp;
        private MainForm _form;

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

                _form.PrepareForExternalProcessing(1255, 1, "Старт создания резервной копии данных");

                string tempFolderPath = GetTempFolderPath();                

                CleanFolder(tempFolderPath);

                IEnumerable<string> notebookIds = GetDistinctNotebooksIds();

                foreach (string id in notebookIds)
                {
                    BackupNotebook(id, tempFolderPath);


                    if (_form.StopExternalProcess)
                        throw new ProcessAbortedByUserException();
                }

                PackfilesToZip(tempFolderPath, filePath);

                CleanFolder(tempFolderPath);                
            }
            catch (ProcessAbortedByUserException)
            {
                BibleCommon.Services.Logger.LogMessage("Process aborted by user");
            }
            finally
            {                
                BibleCommon.Services.Logger.Done();

                _form.ExternalProcessingDone("Создание резервной копии данных успешно завершено.");
            }
        }

        private void PackfilesToZip(string tempFolderPath, string filePath)
        {
            //throw new NotImplementedException();
        }

        private void BackupNotebook(string notebookId, string tempFolderPath)
        {
            _oneNoteApp.Publish(notebookId, Path.Combine(tempFolderPath, 
                OneNoteUtils.GetHierarchyElementName(_oneNoteApp, notebookId)  + ".pckg"), PublishFormat.pfOneNotePackage);
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

        private void CleanFolder(string tempFolderPath)
        {
            foreach (string file in Directory.GetFiles(tempFolderPath))
            {
                File.Delete(file);
            }
        }

        private string GetTempFolderPath()
        {
            string s = Path.Combine(Path.GetDirectoryName(Path.GetDirectoryName(Utils.GetCurrentDirectory())), Consts.TempDirectory);
            if (!Directory.Exists(s))
                Directory.CreateDirectory(s);

            return s;
        }
    }
}
