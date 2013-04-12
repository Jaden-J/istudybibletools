using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.OneNote;
using System.Xml;
using System.Xml.Linq;
using BibleCommon.Helpers;
using System.Xml.XPath;
using BibleCommon.Services;
using BibleCommon.Common;
using BibleCommon.Handlers;
using BibleCommon.Providers;

namespace BibleCommon.Services
{
    public class RelinkAllBibleNotesManager: IDisposable
    {
        private Application _oneNoteApp;
        private NotesPagesProviderManager _notesPagesProviderManager;

        public RelinkAllBibleNotesManager(Application oneNoteApp)
        {
            _oneNoteApp = oneNoteApp;
            _notesPagesProviderManager = new NotesPagesProviderManager();
        }

        public void RelinkBiblePageNotes(string bibleSectionId, string biblePageId, string biblePageName, VersePointer chapterPointer)
        {   
            OneNoteProxy.PageContent biblePageDocument = OneNoteProxy.Instance.GetPageContent(ref _oneNoteApp, biblePageId, OneNoteProxy.PageType.Bible);
            bool wasModified = false;           

            if (OneNoteProxy.Instance.ProcessedVersesOnBiblePagesWithUpdatedLinksToNotesPages.Contains(chapterPointer.ToSimpleVersePointer()))
            {
                var chapterNotesPageLink = NoteLinkManager.GetChapterNotesPageLinkAndCreateIfNeeded(biblePageDocument.Content, biblePageDocument.Xnm);                
                if (RelinkBiblePageNote(bibleSectionId, biblePageId, biblePageName, chapterNotesPageLink, chapterPointer, null))
                    wasModified = true;
            }

            foreach (XElement textElement in biblePageDocument.Content.Root
                .XPathSelectElements("one:Outline/one:OEChildren/one:OE/one:Table/one:Row/one:Cell[2]/one:OEChildren/one:OE/one:T", biblePageDocument.Xnm))
            {
                OneNoteUtils.NormalizeTextElement(textElement);

                XElement bibleVerseElement = textElement.Parent.Parent.Parent.Parent.XPathSelectElement("one:Cell[1]/one:OEChildren/one:OE/one:T", biblePageDocument.Xnm);
                OneNoteUtils.NormalizeTextElement(bibleVerseElement);
                var verseNumber = VerseNumber.GetFromVerseText(bibleVerseElement.Value);

                if (verseNumber.HasValue)
                {
                    VersePointer vp = new VersePointer(chapterPointer, verseNumber.Value.Verse);

                    if (OneNoteProxy.Instance.ProcessedVersesOnBiblePagesWithUpdatedLinksToNotesPages.Contains(vp.ToSimpleVersePointer()))  // если мы обрабатывали этот стих
                    {
                        if (RelinkBiblePageNote(bibleSectionId, biblePageId, biblePageName, textElement, vp, verseNumber))
                            wasModified = true;
                    }
                }
            }

            if (wasModified)
                biblePageDocument.WasModified = true;
        }

        private bool RelinkBiblePageNote(string bibleSectionId, string biblePageId, string biblePageName, XElement textElement, VersePointer vp, VerseNumber? verseNumber)
        {
            bool wasModified = false;
            if (SettingsManager.Instance.StoreNotesPagesInFolder
                // || (SettingsManager.Instance.UseDifferentPagesForEachVerse && SettingsManager.Instance.UseProxyLinksForLinks)   //todo: наверное это не будем делать, сложно будет правильно обрабатывать в OpenNotesPageHandler
                )
            {
                var link = OpenNotesPageHandler.GetCommandUrlStatic(vp, SettingsManager.Instance.ModuleShortName, 
                                                        SettingsManager.Instance.UseDifferentPagesForEachVerse && !vp.IsChapter ? NotesPageType.Verse : NotesPageType.Chapter);                
                string newNotesPageLink = string.Format("<font size='2pt'>{0}</font>",
                                OneNoteUtils.GetLink(SettingsManager.Instance.PageName_Notes, link));

                if (textElement.Value != newNotesPageLink)    //todo: добавить, чтобы это условие срабатывало. То есть правильно енкодить строку
                {
                    textElement.Value = newNotesPageLink;
                    wasModified = true;
                }
            }
            else 
            {
                wasModified = AddOneNoteLinkToNotesPage(bibleSectionId, biblePageId, biblePageName, textElement, verseNumber);
            }

            return wasModified;
        }

        private bool AddOneNoteLinkToNotesPage(string bibleSectionId, string biblePageId, string biblePageName, XElement textElement, VerseNumber? verseNumber)
        {
            bool pageWasCreated;
            string notesPageName = NoteLinkManager.GetDefaultNotesPageName(verseNumber);
            string notesPageId = OneNoteProxy.Instance.GetNotesPageId(ref _oneNoteApp, bibleSectionId, biblePageId, biblePageName, notesPageName, out pageWasCreated);
            string notesRowObjectId = _notesPagesProviderManager.GetNotesRowObjectId(ref _oneNoteApp, notesPageId, verseNumber, !verseNumber.HasValue);

            if (!string.IsNullOrEmpty(notesRowObjectId))
            {
                string newNotesPageLink = string.Format("<font size='2pt'>{0}</font>",
                                    OneNoteUtils.GenerateLink(ref _oneNoteApp, SettingsManager.Instance.PageName_Notes, notesPageId, notesRowObjectId, false));

                textElement.Value = newNotesPageLink;

                return true;
            }

            return false;
        }

        private static bool CantainsLinkToNotesPage(XElement textElement)
        {
            return textElement.Value.IndexOf(string.Format(">{0}<", SettingsManager.Instance.PageName_Notes)) != -1;
        }

        public void Dispose()
        {
            _oneNoteApp = null;
        }
    }
}
