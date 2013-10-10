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
        public string VersesInfoFilePathWithoutExtension { get; set; }

        public AnalyzedVersesService(bool forceUpdate)
        {
            var folder = Utils.GetAnalyzedVersesFolderPath();
            VersesInfoFilePathWithoutExtension = Path.Combine(folder, SettingsManager.Instance.ModuleShortName);

            if (File.Exists(VersesInfoFilePathWithoutExtension + Consts.Constants.FileExtensionCache) && !forceUpdate)
                VersesInfo = SharpSerializationHelper.Deserialize<AnalyzedVersesInfo>(VersesInfoFilePathWithoutExtension + Consts.Constants.FileExtensionCache);
            else
                VersesInfo = new AnalyzedVersesInfo(SettingsManager.Instance.ModuleShortName);
        }       

        public void AddAnalyzedNotebook(string notebookName, string notebookNickname)
        {
            if (!VersesInfo.NotebooksDictionary.ContainsKey(notebookName))
            {
                var maxId = VersesInfo.NotebooksDictionary.Any() ? VersesInfo.NotebooksDictionary.Values.Max(n => n.Id) : 0;
                var notebook = new AnalyzedNotebookInfo() { Name = notebookName, Nickname = notebookNickname, Id = maxId + 1 };

                VersesInfo.NotebooksDictionary.Add(notebookName, notebook);
            }
            else
                VersesInfo.NotebooksDictionary[notebookName].Nickname = notebookNickname;
        }

        public void UpdateVerseInfo(VersePointer verse, decimal weight, decimal detailedWeight, bool isDetailed, string notebookName)
        {
            var verseInfo = VersesInfo
                                .GetOrCreateBookInfo(verse.Book.Index, verse.Book.Name)
                                .GetOrCreateChapterInfo(verse.Chapter.GetValueOrDefault())
                                .GetOrCreateVerseInfo(verse.Verse.GetValueOrDefault());

            if (verseInfo.MaxWeight < weight)
                verseInfo.MaxWeight = weight;

            if (verseInfo.MaxDetailedWeight < detailedWeight)
                verseInfo.MaxDetailedWeight = detailedWeight;

            verseInfo.IsDetailedOnly = verseInfo.IsDetailedOnly && isDetailed;

            var notebookId = VersesInfo.NotebooksDictionary[notebookName].Id;
            if (!verseInfo.Notebooks.Contains(notebookId))
                verseInfo.Notebooks.Add(notebookId);
        }

        public void Update()
        {
            VersesInfo.Sort();
            Utils.SaveToXmlFile(VersesInfo, VersesInfoFilePathWithoutExtension + Consts.Constants.FileExtensionXml);
            SharpSerializationHelper.Serialize(VersesInfo, VersesInfoFilePathWithoutExtension + Consts.Constants.FileExtensionCache);
        }

        public void RemoveContentFiles()
        {
            File.Delete(VersesInfoFilePathWithoutExtension + Consts.Constants.FileExtensionXml);
            File.Delete(VersesInfoFilePathWithoutExtension + Consts.Constants.FileExtensionCache);
        }
    }
}
