using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BibleCommon.Common;
using BibleCommon.Helpers;
using System.IO;

namespace BibleCommon.Services
{
    public class AnalyzedVersesService
    {
        public AnalyzedVersesInfo VersesInfo { get; set; }
        public string VersesInfoFilePath { get; set; }

        public AnalyzedVersesService(bool forceUpdate)
        {
            var folder = Utils.GetAnalyzedVersesFolderPath();
            VersesInfoFilePath = Path.Combine(folder, SettingsManager.Instance.ModuleShortName + ".xml");

            if (File.Exists(VersesInfoFilePath) && !forceUpdate)
                VersesInfo = Utils.LoadFromXmlFile<AnalyzedVersesInfo>(VersesInfoFilePath);
            else
                VersesInfo = new AnalyzedVersesInfo(SettingsManager.Instance.ModuleShortName);
        }

        public void AddAnalyzedNotebook(string notebookName, string notebookNickname)
        {
            var notebook = new AnalyzedNotebookInfo() { Name = notebookName, Nickname = notebookNickname };

            if (!VersesInfo.Notebooks.Contains(notebook))
                VersesInfo.Notebooks.Add(notebook);            
        }

        public void UpdateVerseInfo(VersePointer verse, decimal weight, decimal detailedWeight)
        {
            var verseInfo = VersesInfo
                                .GetOrCreateBookInfo(verse.Book.Index, verse.Book.Name)
                                .GetOrCreateChapterInfo(verse.Chapter.GetValueOrDefault())
                                .GetOrCreateVerseInfo(verse.Verse.GetValueOrDefault());

            if (verseInfo.MaxWeigth < weight)
                verseInfo.MaxWeigth = weight;

            if (verseInfo.MaxDetailedWeigth < detailedWeight)
                verseInfo.MaxDetailedWeigth = detailedWeight;
        }

        public void Update()
        {
            VersesInfo.Sort();
            Utils.SaveToXmlFile(VersesInfo, VersesInfoFilePath);
        }
    }
}
