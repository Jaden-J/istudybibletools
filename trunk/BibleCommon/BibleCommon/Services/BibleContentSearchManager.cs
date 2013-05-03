using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BibleCommon.Common;

namespace BibleCommon.Services
{
    public static class BibleContentSearchManager
    {
        public static bool CheckVerseForExisting(ref VersePointer vp, string pageId, string contentObjectId)
        {
            return GetHierarchyObject(ref vp, pageId, contentObjectId).FoundSuccessfully;
        }

        public static BibleSearchResult GetHierarchyObject(ref VersePointer vp, string pageId, string contentObjectId)
        {
            return GetHierarchyObjectInternal(ref vp, true, pageId, contentObjectId);
        }

        private static BibleSearchResult GetHierarchyObjectInternal(ref VersePointer vp, bool checkForOneChapteredBook, string pageId, string contentObjectId)
        {
            VerseNumber verseNumber;
            BibleSearchResult hierarchySearchResult = null;
            var svp = vp.ToSimpleVersePointer();

            if (SettingsManager.Instance.CurrentBibleContentCached.VerseExists(svp, SettingsManager.Instance.ModuleShortName, out verseNumber))
            {
                hierarchySearchResult = new BibleSearchResult()
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

            if ((checkForOneChapteredBook && vp.IsChapter)
                 && (hierarchySearchResult == null                                               // возможно стих типа "Иуд 2"
                     || ((vp.ParentVersePointer ?? vp).TopChapter.HasValue && vp.Chapter.GetValueOrDefault(0) == 1)))       // Иуд 1-3
            {
                if (SettingsManager.Instance.CurrentBibleContentCached.BookHasOnlyOneChapter(svp))
                {
                    var changedVerseResult = TryToChangeVerseAsOneChapteredBookAndSearchInHierarchy(ref vp, pageId, contentObjectId);
                    if (changedVerseResult != null)
                        return changedVerseResult;
                }
            }            

            if (hierarchySearchResult == null)
            {
                hierarchySearchResult = new BibleSearchResult()
                {
                    ResultType = BibleHierarchySearchResultType.NotFound
                };
                BibleCommon.Services.Logger.LogWarning(pageId, contentObjectId, BibleCommon.Resources.Constants.VerseNotFound, vp.OriginalVerseName);
            }

            return hierarchySearchResult;
        }

        private static BibleSearchResult TryToChangeVerseAsOneChapteredBookAndSearchInHierarchy(ref VersePointer vp, string pageId, string contentObjectId)
        {
            var modifiedVp = new VersePointer(vp.OriginalVerseName);
            modifiedVp.ChangeVerseAsOneChapteredBook();
            var changedVerseResult = GetHierarchyObjectInternal(ref modifiedVp, false, pageId, contentObjectId);
            if (changedVerseResult.FoundSuccessfully)
            {
                vp.ChangeVerseAsOneChapteredBook();
                return changedVerseResult;
            }
            else
                return null;
        }        
    }
}
