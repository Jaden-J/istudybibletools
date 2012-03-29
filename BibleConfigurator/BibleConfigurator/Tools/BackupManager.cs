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

        public void Backup(string folderPath)
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

                _form.PrepareForExternalProcessing(notebookIds.Count(), 1, "Старт создания резервной копии данных");

                folderPath = Path.Combine(folderPath, string.Format("{0}_backup_{1}", Constants.ToolsName, DateTime.Now.ToShortDateString()));

                if (!Directory.Exists(folderPath))
                    Directory.CreateDirectory(folderPath);

                

                foreach (string id in notebookIds)
                {
                    string notebookName = OneNoteUtils.GetHierarchyElementName(_oneNoteApp, id);
                    _form.PerformProgressStep(string.Format("Старт создания резервной копии записной книжки '{0}'", notebookName));

                    BackupNotebook(id, notebookName, folderPath);                    
                    
                    if (_form.StopExternalProcess)
                        throw new ProcessAbortedByUserException();
                }                
            }
            catch (ProcessAbortedByUserException)
            {
                BibleCommon.Services.Logger.LogMessage("Process aborted by user");
            }
            finally
            {                
                BibleCommon.Services.Logger.Done();

                _form.ExternalProcessingDone("Создание резервной копии данных инициировано. Операция займёт несколько минут. Данную программу можно закрыть.");
            }
        }        

        private void BackupNotebook(string notebookId, string notebookName, string tempFolderPath)
        {
            _oneNoteApp.Publish(notebookId, Path.Combine(tempFolderPath,
                notebookName + ".onepkg"), PublishFormat.pfOneNotePackage);
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
    }
}
