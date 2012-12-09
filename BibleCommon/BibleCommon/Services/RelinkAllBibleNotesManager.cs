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

namespace BibleCommon.Services
{
    public class RelinkAllBibleNotesManager: IDisposable
    {
        private Application _oneNoteApp;        

        public RelinkAllBibleNotesManager(Application oneNoteApp)
        {
            _oneNoteApp = oneNoteApp;            
        }

        public void RelinkBiblePageNotes(string bibleSectionId, string biblePageId, string biblePageName, VersePointer chapterPointer)
        {   
            OneNoteProxy.PageContent biblePageDocument = OneNoteProxy.Instance.GetPageContent(ref _oneNoteApp, biblePageId, OneNoteProxy.PageType.Bible);
            bool wasModified = false;

            XElement chapterNotesPageLink = NoteLinkManager.GetChapterNotesPageLink(biblePageDocument.Content, biblePageDocument.Xnm);

            if (chapterNotesPageLink != null)
                if (RelinkBiblePageNote(bibleSectionId, biblePageId, biblePageName, chapterNotesPageLink, null))
                    wasModified = true;

            foreach (XElement textElement in biblePageDocument.Content.Root
                .XPathSelectElements("one:Outline/one:OEChildren/one:OE/one:Table/one:Row/one:Cell[2]/one:OEChildren/one:OE/one:T", biblePageDocument.Xnm))
            {
                if (!string.IsNullOrEmpty(textElement.Value))
                {
                    OneNoteUtils.NormalizeTextElement(textElement);

                    if (CantainsLinkToNotesPage(textElement))
                    {
                        XElement bibleVerseElement = textElement.Parent.Parent.Parent.Parent.XPathSelectElement("one:Cell[1]/one:OEChildren/one:OE/one:T", biblePageDocument.Xnm);
                        OneNoteUtils.NormalizeTextElement(bibleVerseElement);
                        var verseNumber = VerseNumber.GetFromVerseText(bibleVerseElement.Value);

                        if (verseNumber.HasValue)
                        {
                            VersePointer vp = new VersePointer(chapterPointer, verseNumber.Value.Verse);
                            
                            if (OneNoteProxy.Instance.ProcessedVersesOnBiblePagesWithUpdatedLinksToNotesPages.Contains(vp.ToSimpleVersePointer()))  // если мы обрабатывали этот стих
                            {
                                if (RelinkBiblePageNote(bibleSectionId, biblePageId, biblePageName, textElement, verseNumber))
                                    wasModified = true;
                            }
                        }
                    }
                }
            }

            if (wasModified)
                biblePageDocument.WasModified = true;
        }

        private bool RelinkBiblePageNote(string bibleSectionId, string biblePageId, string biblePageName, XElement textElement, VerseNumber? verseNumber)
        {
            bool pageWasCreated;
            string notesPageName = NoteLinkManager.GetDefaultNotesPageName(verseNumber);
            string notesPageId = OneNoteProxy.Instance.GetNotesPageId(ref _oneNoteApp, bibleSectionId, biblePageId, biblePageName, notesPageName, out pageWasCreated);
            string notesRowObjectId = NotesPageManager.GetNotesRowObjectId(ref _oneNoteApp, notesPageId, verseNumber, !verseNumber.HasValue);

            if (!string.IsNullOrEmpty(notesRowObjectId))
            {
                string newNotesPageLink = string.Format("<font size='2pt'>{0}</font>",
                                    OneNoteUtils.GenerateHref(ref _oneNoteApp, SettingsManager.Instance.PageName_Notes, notesPageId, notesRowObjectId));

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
