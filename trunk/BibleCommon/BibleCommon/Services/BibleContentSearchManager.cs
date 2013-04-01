using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BibleCommon.Common;

namespace BibleCommon.Services
{
    public static class BibleContentSearchManager
    {
        public static HierarchySearchManager.HierarchySearchResult GetHierarchyObject(VersePointer vp)
        {
            VerseNumber verseNumber;
            HierarchySearchManager.HierarchySearchResult hierarchySearchResult = null;

            if (SettingsManager.Instance.CurrentBibleContentCached.VerseExists(vp.ToSimpleVersePointer(), SettingsManager.Instance.ModuleShortName, out verseNumber))
            {
                hierarchySearchResult = new HierarchySearchManager.HierarchySearchResult()
                {
                    ResultType = BibleHierarchySearchResultType.Successfully,
                    HierarchyStage = vp.IsChapter
                                        ? BibleHierarchyStage.Page
                                        : BibleHierarchyStage.ContentPlaceholder,
                    HierarchyObjectInfo = new BibleHierarchyObjectInfo()
                    {
                        VerseNumber = verseNumber
                    }
                };
            }
            else if (vp.IsChapter)   // возможно стих типа "2Ин 8"
            {
                //if (BookHasOnlyOneChapter(ref oneNoteApp, vp, result.HierarchyObjectInfo, useCacheIfAvailable))
                //{
                //    var changedVerseResult = TryToChangeVerseAsOneChapteredBookAndSearchInHierarchy(
                //        ref oneNoteApp, bibleNotebookId, ref vp, findAllVerseObjects, useCacheIfAvailable, false);
                //    if (changedVerseResult != null)
                //        return changedVerseResult;
                //}
            }
            

            if (hierarchySearchResult == null)
            {
                hierarchySearchResult = new HierarchySearchManager.HierarchySearchResult()
                {
                    ResultType = BibleHierarchySearchResultType.NotFound
                };
                BibleCommon.Services.Logger.LogWarning(BibleCommon.Resources.Constants.VerseNotFound, vp.OriginalVerseName);
            }

            return hierarchySearchResult;
        }
    }
}
