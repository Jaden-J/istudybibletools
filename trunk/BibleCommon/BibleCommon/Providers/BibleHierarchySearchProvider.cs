using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BibleCommon.Common;
using Microsoft.Office.Interop.OneNote;
using BibleCommon.Services;

namespace BibleCommon.Providers
{
    public static class BibleHierarchySearchProvider
    {
        public static BibleSearchResult GetHierarchyObject(ref Application oneNoteApp, ref VersePointer vp, NoteLinkManager.AnalyzeDepth? linkDepth, string pageId, string contentObjectId)
        {
            if (SettingsManager.Instance.UseProxyLinksForBibleVerses && SettingsManager.Instance.CanUseBibleContent
                && (SettingsManager.Instance.StoreNotesPagesInFolder || linkDepth.GetValueOrDefault() < NoteLinkManager.AnalyzeDepth.Full))
            {
                return BibleContentSearchManager.GetHierarchyObject(ref vp, pageId, contentObjectId);
            }
            else
            {
                return HierarchySearchManager.GetHierarchyObject(
                                                    ref oneNoteApp, SettingsManager.Instance.NotebookId_Bible, ref vp, HierarchySearchManager.FindVerseLevel.AllVerses, pageId, contentObjectId);
            }
        }

        public static bool CheckVerseForExisting(ref Application oneNoteApp, ref VersePointer vp, string pageId, string contentObjectId)
        {
            if (SettingsManager.Instance.CanUseBibleContent)
            {
                return BibleContentSearchManager.GetHierarchyObject(ref vp, pageId, contentObjectId).FoundSuccessfully;
            }
            else
            {
                return HierarchySearchManager.GetHierarchyObject(
                                                    ref oneNoteApp, SettingsManager.Instance.NotebookId_Bible, ref vp,
                                                    HierarchySearchManager.FindVerseLevel.OnlyVersesOfFirstChapter, pageId, contentObjectId).FoundSuccessfully;
            }
        }
    }
}
