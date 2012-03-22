using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BibleCommon.Services;
using BibleNoteLinkerEx.Properties;

namespace BibleNoteLinkerEx
{
    public static class Helper
    {
        public static List<string> GetSelectedNotebooksIds()
        {
            if (!string.IsNullOrEmpty(Settings.Default.SelectedNotebooks))
            {
                return Settings.Default.SelectedNotebooks.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries).ToList();
            }
            else
            {
                if (SettingsManager.Instance.IsSingleNotebook)
                {
                    return new List<string>() 
                    {
                        SettingsManager.Instance.SectionGroupId_BibleStudy, 
                        SettingsManager.Instance.SectionGroupId_BibleComments
                    };
                }
                else
                {
                    return new List<string>() 
                    {
                        SettingsManager.Instance.NotebookId_BibleStudy, 
                        SettingsManager.Instance.NotebookId_BibleComments
                    };
                }
            }
        }

        public static void SaveSelectedNotebooksIds(List<string> notebooksIds)
        {
            Settings.Default.SelectedNotebooks = string.Join(";", notebooksIds.ToArray());
            Settings.Default.Save();
        }
    }
}
