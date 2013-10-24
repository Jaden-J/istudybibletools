using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using BibleCommon.Common;
using BibleCommon.Helpers;
using BibleCommon.UI.Forms;
using Microsoft.Office.Interop.OneNote;

namespace BibleCommon.Services
{
    public class UpdateManager : IDisposable
    {
        private Application _oneNoteApp;
        public Application OneNoteApp
        {
            get
            {
                if (_oneNoteApp == null)
                    _oneNoteApp = OneNoteUtils.CreateOneNoteAppSafe();

                return _oneNoteApp;                
            }
        }

        public void TryToApplyUpdateCommands()
        {
            if (SettingsManager.Instance.SettingsWereLoadedFromFile)        // то есть если новая установка - то не надо применять обновления
            {
                bool notesWasRegenerated = false;
                if (SettingsManager.Instance.VersionFromSettings == null || SettingsManager.Instance.VersionFromSettings < new Version(3, 1))
                {
                    Utils.DoWithExceptionHandling(false, () => TryToRegenerateNotesPages(true));
                    Utils.DoWithExceptionHandling(false, BibleParallelTranslationManager.MergeAllModulesWithMainBible);  // так как раньше оно не правильно вызывалось и вызывалось ли вообще...

                    notesWasRegenerated = true;
                }


                if (SettingsManager.Instance.VersionFromSettings == null || SettingsManager.Instance.VersionFromSettings < new Version(3, 2))
                {
                    Utils.DoWithExceptionHandling(false, () =>
                    {
                        if (ApplicationCache.Instance.IsBibleVersesLinksCacheActive)
                            ApplicationCache.Instance.CleanBibleVersesLinksCache(false);
                    });
                }

                if (SettingsManager.Instance.VersionFromSettings == null || SettingsManager.Instance.VersionFromSettings < new Version(3, 2, 6))
                {
                    if (!notesWasRegenerated)
                        Utils.DoWithExceptionHandling(false, () => TryToRegenerateNotesPages(false));
                }

                if (SettingsManager.Instance.VersionFromSettings == null || SettingsManager.Instance.VersionFromSettings < new Version(3, 2, 16))
                {
                    Utils.DoWithExceptionHandling(false, NotesPageManagerFS.UpdateResources);
                }
            }
        }


        private void TryToRegenerateNotesPages(bool writeNotesPages)
        {
            if (!string.IsNullOrEmpty(SettingsManager.Instance.FolderPath_BibleNotesPages) && !string.IsNullOrEmpty(SettingsManager.Instance.ModuleShortName))
            {
                var service = new AnalyzedVersesService(true);

                AddDefaultAnalyzedNotebooksInfo(service);
                RegenerateNotesPages(service, writeNotesPages);

                service.Update();
                NotesPageManagerFS.UpdateResources();
            }
        }

        private void RegenerateNotesPages(AnalyzedVersesService service, bool writeNotesPages)
        {
            if (Directory.Exists(SettingsManager.Instance.FolderPath_BibleNotesPages))
            {
                var files = Directory.GetFiles(SettingsManager.Instance.FolderPath_BibleNotesPages, "*.htm", SearchOption.AllDirectories);
                using (var form = new ProgressForm(BibleCommon.Resources.Constants.UpgradingNotesPages, false, (f) =>
                {
                    foreach (var filePath in files)
                    {
                        Utils.DoWithExceptionHandling(true, () =>
                        {
                            var fileContent = File.ReadAllText(filePath);
                            var startTitleIndex = fileContent.IndexOf("<title>") + "<title>".Length;
                            if (startTitleIndex > 10)
                            {
                                var endTitleIndex = fileContent.IndexOf("</title>");
                                var title = fileContent.Substring(startTitleIndex, endTitleIndex - startTitleIndex);
                                var parts = title.Split(new char[] { '[', ']' }, StringSplitOptions.RemoveEmptyEntries);
                                var pageName = parts[0];
                                var chapterPointer = new VersePointer(parts[1]);
                                var pageData = new NotesPageData(filePath, pageName, Path.GetFileNameWithoutExtension(filePath) == "0" ? NotesPageType.Chapter : NotesPageType.Verse, chapterPointer, true);
                                pageData.Serialize(ref _oneNoteApp, service, writeNotesPages);
                            }

                            f.PerformStep(BibleCommon.Resources.Constants.UpgradingFile + ": ...\\" + Path.Combine(
                                                                    Path.GetFileName(Path.GetDirectoryName(Path.GetDirectoryName(filePath))),
                                                                    Path.Combine(
                                                                        Path.GetFileName(Path.GetDirectoryName(filePath)),
                                                                        Path.GetFileName(filePath))));
                        });
                    }
                })
                    )
                {
                    form.ShowDialog(files.Count());
                }
            }
        }

        private void AddDefaultAnalyzedNotebooksInfo(AnalyzedVersesService service)
        {
            if (SettingsManager.Instance.SelectedNotebooksForAnalyze != null)
            {
                foreach (var notebookInfo in SettingsManager.Instance.SelectedNotebooksForAnalyze)
                {
                    Utils.DoWithExceptionHandling(true, () =>
                    {
                        var notebookName = OneNoteUtils.GetHierarchyElementName(ref _oneNoteApp, notebookInfo.NotebookId);
                        var notebookNickname = OneNoteUtils.GetNotebookElementNickname(ref _oneNoteApp, notebookInfo.NotebookId);

                        service.AddAnalyzedNotebook(notebookName, notebookNickname);
                    });
                }
            }
        }

        public void Dispose()
        {
            OneNoteUtils.ReleaseOneNoteApp(ref _oneNoteApp);
        }
    }
}
