using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BibleCommon.Services;
using BibleNoteLinker.Properties;

namespace BibleNoteLinker
{
    public static class Helper
    {
        public static List<string> GetSelectedNotebooksIds()
        {
            if (!string.IsNullOrEmpty(Settings.Default.SelectedNotebooks))
            {
                string[] s = Settings.Default.SelectedNotebooks.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);

                if (s.Length == 2 && s[0] == SettingsManager.Instance.IsSingleNotebook.ToString().ToLower())
                {
                    return s[1].Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                }
                else
                {
                    SaveSelectedNotebooksIds(string.Empty);
                    return GetSelectedNotebooksIds();
                }                
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

        private static void SaveSelectedNotebooksIds(string notebookIds)
        {
            Settings.Default.SelectedNotebooks = notebookIds;
            Settings.Default.Save();
        }


        public static void SaveSelectedNotebooksIds(List<string> notebooksIds)
        {
            SaveSelectedNotebooksIds(string.Join(";", new string[] 
            {
                SettingsManager.Instance.IsSingleNotebook.ToString().ToLower(), string.Join(",", notebooksIds.ToArray())
            }));            
        }

        public static string GetRightFoundPagesString(int pagesCount)
        {
            string firstPart = pagesCount == 1 ? BibleCommon.Resources.Constants.NoteLinkerOneFound : BibleCommon.Resources.Constants.NoteLinkerManyFound;

            return string.Format("{0} {1}", firstPart, GetRightPagesString(pagesCount));
        }


        public static string GetRightPagesString(int pagesCount)
        {
            string s = BibleCommon.Resources.Constants.NoteLinkerOfManyPages;
            int tempPagesCount = pagesCount;

            tempPagesCount = tempPagesCount % 100;
            if (!(tempPagesCount >= 10 && tempPagesCount <= 20))
            {
                tempPagesCount = tempPagesCount % 10;

                if (tempPagesCount == 1)
                    s = BibleCommon.Resources.Constants.NoteLinkerOnePage;
                else if (tempPagesCount >= 2 && tempPagesCount <= 4)
                    s = BibleCommon.Resources.Constants.NoteLinkerManyPages;
            }            

            return string.Format("{0} {1}", pagesCount, s);
        }

    }
}
