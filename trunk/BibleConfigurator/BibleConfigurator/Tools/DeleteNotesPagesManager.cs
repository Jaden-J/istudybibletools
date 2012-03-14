using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.OneNote;
using BibleCommon.Helpers;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;
using BibleCommon.Services;
using BibleCommon.Consts;
using BibleCommon.Common;

namespace BibleConfigurator.Tools
{
    public class DeleteNotesPagesManager
    {
        private Application _oneNoteApp;
        private MainForm _form;

        public DeleteNotesPagesManager(Application oneNoteApp, MainForm form)
        {
            _oneNoteApp = oneNoteApp;
            _form = form;
        }

        public void DeleteNotesPages()
        {
            if (!SettingsManager.Instance.IsConfigured(_oneNoteApp))
            {
                Logger.LogError(Constants.Error_SystemIsNotConfigures);
                return;
            }

            try
            {
                Dictionary<string, string> pagesToDelete = GetAllNotesPagesIds();

                _form.PrepareForExternalProcessing(1255 + pagesToDelete.Count, 1, "Старт удаления страниц 'Сводные заметок'.");

                new NotebookIterator(_oneNoteApp).Iterate("DeleteNotesPagesManager",
                    SettingsManager.Instance.NotebookId_Bible, SettingsManager.Instance.SectionGroupId_Bible, pageInfo =>
                        {
                            _form.PerformProgressStep(string.Format("Обработка страницы '{0}'",  pageInfo.PageName));         

                            try
                            {
                                DeleteAllNotesOnPage(pageInfo.SectionGroupId, pageInfo.SectionId, pageInfo.PageId, pageInfo.PageName);
                            }
                            catch (Exception ex)
                            {
                                Logger.LogError(ex.ToString());
                            }

                            if (_form.StopExternalProcess)
                                throw new ProcessAbortedByUserException();
                        });

                foreach (var page in pagesToDelete)
                {
                    string message = string.Format("Удаление страницы '{0}'", page.Value);
                    _form.PerformProgressStep(message);
                    BibleCommon.Services.Logger.LogMessage(message);

                    DeleteNotesPage(page.Key);

                    if (_form.StopExternalProcess)
                        throw new ProcessAbortedByUserException();
                }
            }
            catch (ProcessAbortedByUserException)
            {
                BibleCommon.Services.Logger.LogMessage("Process aborted by user");
            }
            finally
            {
                BibleCommon.Services.Logger.Done();

                _form.ExternalProcessingDone("Обновление ссылок на комментарии успешно завершено");
            }
        }

        private Dictionary<string, string> GetAllNotesPagesIds()
        {
            Dictionary<string, string> result = new Dictionary<string, string>();

            var allPages = OneNoteProxy.Instance.GetHierarchy(_oneNoteApp, SettingsManager.Instance.NotebookId_BibleNotesPages, HierarchyScope.hsPages, true);

            foreach(var page in allPages.Content.XPathSelectElements("//one:Page", allPages.Xnm))
            {
                if (!OneNoteUtils.IsRecycleBin(page))
                {
                    string pageName = (string)page.Attribute("name");
                    bool isSummaryNotesPage = false;

                    var metaData = page.XPathSelectElement("one:Meta", allPages.Xnm);
                    if (metaData != null)
                    {
                        var name = (string)metaData.Attribute("name");
                        var content = (string)metaData.Attribute("content");

                        if (name == Constants.Key_IsSummaryNotesPage && bool.Parse(content))
                            isSummaryNotesPage = true;
                    }

                    if (!isSummaryNotesPage)
                    {
                        // for back compatibility                    
                        if (pageName.StartsWith(SettingsManager.Instance.PageName_Notes + ".")
                            || pageName.StartsWith(SettingsManager.Instance.PageName_RubbishNotes + "."))
                            isSummaryNotesPage = true;
                    }

                    if (isSummaryNotesPage)
                        result.Add((string)page.Attribute("ID"), pageName);
                }
            }

            return result;
        }

        private void DeleteAllNotesOnPage(string bibleSectionGroupId, string bibleSectionId, string biblePageId, string biblePageName)
        {   
            bool wasModified = false;
            string pageContentXml;
            XDocument notePageDocument;
            XmlNamespaceManager xnm;
            _oneNoteApp.GetPageContent(biblePageId, out pageContentXml);
            notePageDocument = OneNoteUtils.GetXDocument(pageContentXml, out xnm);

            foreach (XElement noteTextElement in notePageDocument.Root.XPathSelectElements("//one:Table/one:Row/one:Cell[2]/one:OEChildren/one:OE/one:T", xnm))
            {
                if (!string.IsNullOrEmpty(noteTextElement.Value))
                {
                    if (CantainsLinkToNotesPage(noteTextElement))
                    {
                        noteTextElement.Value = string.Empty;
                        wasModified = true;
                    }
                }
            }

            XElement chapterNotesLink = FindChapterNotesLink(notePageDocument, xnm);
            if (chapterNotesLink != null)
            {
                _oneNoteApp.DeletePageContent(biblePageId, (string)chapterNotesLink.Attribute("objectID"));
                chapterNotesLink.Remove();
                wasModified = true;
            }            

            if (wasModified)
                OneNoteUtils.UpdatePageContentSafe(_oneNoteApp, notePageDocument);
        }

        private void DeleteNotesPage(string notesPageId)
        {
            if (!string.IsNullOrEmpty(notesPageId))
            {
                string sectionId;
                _oneNoteApp.GetHierarchyParent(notesPageId, out sectionId);

                _oneNoteApp.DeleteHierarchy(notesPageId);

                string sectionPagesXml;
                XmlNamespaceManager xnm;
                _oneNoteApp.GetHierarchy(sectionId, HierarchyScope.hsPages, out sectionPagesXml);
                XDocument sectionPages = OneNoteUtils.GetXDocument(sectionPagesXml, out xnm);
                if (sectionPages.Root.XPathSelectElements("one:Page", xnm).Count() == 0)
                    _oneNoteApp.DeleteHierarchy(sectionId);  // удаляем раздел, если нет больше в нём страниц
            }
        }

        private static XElement FindChapterNotesLink(XDocument notePageDocument, XmlNamespaceManager xnm)
        {
            foreach (XElement outline in notePageDocument.Root.XPathSelectElements("//one:Outline", xnm))
            {
                List<XElement> textElements = outline.XPathSelectElements(".//one:T", xnm).ToList();
                if (textElements.Count == 1)
                {
                    if (CantainsLinkToNotesPage(textElements.First()))
                    {
                        return outline;
                    }
                }
            }

            return null;
        }

        private static bool CantainsLinkToNotesPage(XElement textElement)
        {
            return textElement.Value.IndexOf(string.Format(">{0}<", SettingsManager.Instance.PageName_Notes)) != -1;
        }
    }
}
