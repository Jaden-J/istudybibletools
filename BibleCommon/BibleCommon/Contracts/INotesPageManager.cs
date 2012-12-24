using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.OneNote;
using BibleCommon.Services;
using BibleCommon.Common;

namespace BibleCommon.Contracts
{
    public interface INotesPageManager
    {
        string ManagerName { get; }        

        string UpdateNotesPage(ref Application oneNoteApp, NoteLinkManager noteLinkManager, VersePointer vp, int versePointeHtmlStartIndex, bool isChapter,
           HierarchySearchManager.HierarchyObjectInfo verseHierarchyObjectInfo,
           HierarchyElementInfo notePageId, string notesPageId, string notePageContentObjectId,
           string notesPageName, int notesPageWidth, bool force, bool processAsExtendedVerse, bool commonNotesPage, out bool rowWasAdded);

        string GetNotesRowObjectId(ref Application oneNoteApp, string notesPageId, VerseNumber? verseNumber, bool isChapter);
    }
}
