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
    public class DeleteNotesPagesManager: IDisposable
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
            if (!SettingsManager.Instance.IsConfigured(ref _oneNoteApp))
            {
                FormLogger.LogError(BibleCommon.Resources.Constants.Error_SystemIsNotConfigured);
                return;
            }

            try
            {
                BibleCommon.Services.Logger.Init("DeleteNotesPagesManager");

                Dictionary<string, string> pagesToDelete = GetAllNotesPagesIds();
                int chaptersCount = ModulesManager.GetBibleChaptersCount(SettingsManager.Instance.ModuleShortName, true);

                _form.PrepareForLongProcessing(chaptersCount + pagesToDelete.Count, 1, BibleCommon.Resources.Constants.DeleteNotesPagesManagerStartMessage);

                NotebookIteratorHelper.Iterate(ref _oneNoteApp,
                    SettingsManager.Instance.NotebookId_Bible, SettingsManager.Instance.SectionGroupId_Bible, pageInfo =>
                        {
                            _form.PerformProgressStep(string.Format("{0} '{1}'", BibleCommon.Resources.Constants.ProcessPage , pageInfo.Title));

                            try
                            {
                                DeleteAllNotesOnPage(pageInfo.SectionGroupId, pageInfo.SectionId, pageInfo.Id, pageInfo.Title);
                            }
                            catch (Exception ex)
                            {
                                FormLogger.LogError(ex.ToString());
                            }

                            if (_form.StopLongProcess)
                                throw new ProcessAbortedByUserException();
                        });

                foreach (var page in pagesToDelete)
                {
                    string message = string.Format("{0} '{1}'", BibleCommon.Resources.Constants.DeleteNotesPagesManagerRemovePage, page.Value);
                    _form.PerformProgressStep(message);
                    BibleCommon.Services.Logger.LogMessage(message);

                    DeleteNotesPage(page.Key);

                    if (_form.StopLongProcess)
                        throw new ProcessAbortedByUserException();
                }

                _form.LongProcessingDone(BibleCommon.Resources.Constants.DeleteNotesPagesManagerFinishMessage);
            }
            catch (ProcessAbortedByUserException)
            {
                BibleCommon.Services.Logger.LogMessage("Process aborted by user");
            }
            finally
            {
                BibleCommon.Services.Logger.Done();               
            }
        }

        private Dictionary<string, string> GetAllNotesPagesIds()
        {
            Dictionary<string, string> result = new Dictionary<string, string>();

            var allPages = OneNoteProxy.Instance.GetHierarchy(ref _oneNoteApp, SettingsManager.Instance.NotebookId_BibleNotesPages, HierarchyScope.hsPages, true);

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
            string pageContentXml = null;
            XDocument notePageDocument;
            XmlNamespaceManager xnm;

            OneNoteUtils.UseOneNoteAPI(ref _oneNoteApp, (oneNoteAppSafe) =>
            {
                oneNoteAppSafe.GetPageContent(biblePageId, out pageContentXml, PageInfo.piBasic, Constants.CurrentOneNoteSchema);
            });

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
                OneNoteUtils.UseOneNoteAPI(ref _oneNoteApp, (oneNoteAppSafe) =>
                {
                    oneNoteAppSafe.DeletePageContent(biblePageId, (string)chapterNotesLink.Attribute("objectID"));
                });
                chapterNotesLink.Remove();
                wasModified = true;
            }            

            if (wasModified)
                OneNoteUtils.UpdatePageContentSafe(ref _oneNoteApp, notePageDocument, xnm);
        }

        private void DeleteNotesPage(string notesPageId)
        {
            if (!string.IsNullOrEmpty(notesPageId))
            {
                string sectionId = null;
                OneNoteUtils.UseOneNoteAPI(ref _oneNoteApp, (oneNoteAppSafe) =>
                {
                    oneNoteAppSafe.GetHierarchyParent(notesPageId, out sectionId);
                });

                OneNoteUtils.UseOneNoteAPI(ref _oneNoteApp, (oneNoteAppSafe) =>
                {
                    oneNoteAppSafe.DeleteHierarchy(notesPageId, DateTime.MinValue, true);
                });

                string sectionPagesXml = null;
                XmlNamespaceManager xnm;

                OneNoteUtils.UseOneNoteAPI(ref _oneNoteApp, (oneNoteAppSafe) =>
                {
                    oneNoteAppSafe.GetHierarchy(sectionId, HierarchyScope.hsPages, out sectionPagesXml, Constants.CurrentOneNoteSchema);
                });

                XDocument sectionPages = OneNoteUtils.GetXDocument(sectionPagesXml, out xnm);
                if (sectionPages.Root.XPathSelectElements("one:Page", xnm).Count() == 0)
                {
                    OneNoteUtils.UseOneNoteAPI(ref _oneNoteApp, (oneNoteAppSafe) =>
                    {
                        oneNoteAppSafe.DeleteHierarchy(sectionId, DateTime.MinValue, true);  // удаляем раздел, если нет больше в нём страниц
                    });
                }
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

        public void Dispose()
        {
            _oneNoteApp = null;
            _form = null;
        }
    }
}
