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
    public class RelinkAllBibleCommentsManager: IDisposable
    {
        private Application _oneNoteApp;
        private MainForm _form;

        public RelinkAllBibleCommentsManager(MainForm form)
        {
            _oneNoteApp = OneNoteUtils.CreateOneNoteAppSafe();
            _form = form;
        }

        public void RelinkAllBibleComments()
        {
            if (!SettingsManager.Instance.IsConfigured(ref _oneNoteApp))
            {
                FormLogger.LogError(BibleCommon.Resources.Constants.Error_SystemIsNotConfigured);
                return;
            }   

            try
            {
                BibleCommon.Services.Logger.Init("RelinkAllBibleCommentsManager");

                try
                {
                    OneNoteLocker.UnlockBible(ref _oneNoteApp, true, () => _form.StopLongProcess);
                }
                catch (NotSupportedException)
                {
                    //todo: log it
                }

                int chaptersCount = ModulesManager.GetBibleChaptersCount(SettingsManager.Instance.ModuleShortName, true);
                _form.PrepareForLongProcessing(chaptersCount, 1, BibleCommon.Resources.Constants.RelinkBibleCommentsManagerStartMessage);                

                NotebookIteratorHelper.Iterate(ref _oneNoteApp,
                    SettingsManager.Instance.NotebookId_Bible, SettingsManager.Instance.SectionGroupId_Bible, pageInfo =>
                        {
                            try
                            {
                                RelinkPageComments(pageInfo.SectionId, pageInfo.Id, pageInfo.Title);
                            }
                            catch (Exception ex)
                            {
                                FormLogger.LogError(ex.ToString());
                            }

                            if (_form.StopLongProcess)
                                throw new ProcessAbortedByUserException();
                        });

                _form.LongProcessingDone(BibleCommon.Resources.Constants.RelinkBibleCommentsManagerFinishMessage);
            }
            catch (ProcessAbortedByUserException)
            {
                BibleCommon.Services.Logger.LogMessageParams("Process aborted by user");
            }
            finally
            {                
                BibleCommon.Services.Logger.Done();                
            }
        }

        private void RelinkPageComments(string sectionId, string pageId, string pageName)
        {
            _form.PerformProgressStep(string.Format("{0} '{1}'", BibleCommon.Resources.Constants.ProcessPage, pageName));

            string pageContent = null;
            XmlNamespaceManager xnm;
            OneNoteUtils.UseOneNoteAPI(ref _oneNoteApp, () =>
            {
                _oneNoteApp.GetPageContent(pageId, out pageContent, PageInfo.piBasic, Constants.CurrentOneNoteSchema);
            });
            XDocument pageDocument = OneNoteUtils.GetXDocument(pageContent, out xnm);            
            bool wasModified = false;

            foreach (XElement textElement in pageDocument.Root.XPathSelectElements("//one:Outline/one:OEChildren/one:OE/one:Table/one:Row/one:Cell[1]/one:OEChildren/one:OE/one:T", xnm))
            {
                OneNoteUtils.NormalizeTextElement(textElement);

                int linkIndex = textElement.Value.IndexOf("<a ");

                while (linkIndex > -1)
                {
                    int linkEnd = textElement.Value.IndexOf("</a>", linkIndex + 1);

                    if (linkEnd != -1)
                    {
                        if (RelinkPageComment(sectionId, pageId, pageName, textElement, linkIndex, linkEnd))                          
                            wasModified = true;                        
                    }

                    linkIndex = textElement.Value.IndexOf("<a ", linkIndex + 1);
                }                
            }

            if (wasModified)
                OneNoteUtils.UpdatePageContentSafe(ref _oneNoteApp, pageDocument, xnm);
        }

        private bool RelinkPageComment(string bibleSectionId, string biblePageId, string biblePageName, XElement textElement, int linkIndex, int linkEnd)
        {
            string commentLink = textElement.Value.Substring(linkIndex, linkEnd - linkIndex + "</a>".Length);
            string commentText = StringUtils.GetText(commentLink);

            string commentPageName = GetCommentPageName(commentLink);
            bool pageWasCreated;
            string commentPageId = ApplicationCache.Instance.GetCommentPageId(ref _oneNoteApp, bibleSectionId, biblePageId, biblePageName, commentPageName, out pageWasCreated, false);
            if (!string.IsNullOrEmpty(commentPageId))
            {
                string commentObjectId = GetComentObjectId(commentPageId, commentText, null, 0);

                    string newCommentLink = OneNoteUtils.GenerateLink(ref _oneNoteApp, commentText, commentPageId, commentObjectId);

                    textElement.Value = textElement.Value.Replace(commentLink, newCommentLink);

                    return true;
                }

            return false;
        }

        private string GetComentObjectId(string commentPageId, string commentText, string textElementId, int startIndex)
        {            
            ApplicationCache.PageContent pageDoc = ApplicationCache.Instance.GetPageContent(ref _oneNoteApp, commentPageId, ApplicationCache.PageType.CommentPage);

            foreach (XElement el in pageDoc.Content.Root.XPathSelectElements(string.Format("one:Outline/one:OEChildren/one:OE{0}/one:T",
                !string.IsNullOrEmpty(textElementId) ? string.Format("[@objectID=\"{0}\"]", textElementId) : string.Empty), pageDoc.Xnm))
            {
                OneNoteUtils.NormalizeTextElement(el);

                bool needToSearchVerse = true;
                if (el.Value.Length > startIndex + 1)
                {
                    int boldTagIndex = el.Value.IndexOf("font-weight:bold", startIndex + 1);
                    if (boldTagIndex != -1)
                    {
                        boldTagIndex = el.Value.IndexOf(">", boldTagIndex + 1);

                        if (boldTagIndex != -1)
                        {
                            int textBreakIndex;
                            int htmlBreakIndex;
                            string textBefore = StringUtils.GetPrevString(el.Value, boldTagIndex + 1, new SearchMissInfo(boldTagIndex, SearchMissInfo.MissMode.CancelOnNextMiss),
                                    out textBreakIndex, out htmlBreakIndex, StringSearchIgnorance.None, StringSearchMode.NotSpecified).Replace("&nbsp;", "");

                            if (textBefore.Trim().Length <= 5)  // чтоб убедиться, что мы взяли текст в начале строки
                            {
                                int boldEndIndex = el.Value.IndexOf("</span>", boldTagIndex + 1);

                                if (boldEndIndex != -1)
                                {
                                    string commentValue = el.Value.Substring(boldTagIndex + 1, boldEndIndex - boldTagIndex - 1);

                                    if (!string.IsNullOrEmpty(commentValue))
                                    {
                                        string objectId = (string)el.Parent.Attribute("objectID");

                                        if (string.IsNullOrEmpty(commentValue.Trim()))
                                        {
                                            string nextCommentObjectId = GetComentObjectId(commentPageId, commentText, objectId, boldEndIndex);
                                            if (!string.IsNullOrEmpty(nextCommentObjectId))
                                                return nextCommentObjectId;
                                        }

                                        if (commentValue == commentText)
                                            return objectId;
                                        else
                                            needToSearchVerse = false;  // это точно не стих, это просто другой комментарий
                                    }
                                }
                            }
                        }
                    }
                }

                if (needToSearchVerse)
                {
                    // если дошли до сюда, значит не нашли там
                    int temp;
                    if (int.TryParse(commentText, out temp))  // значит скорее всего указали стих
                    {
                        string verseStartSearchString = ">:";
                        int verseStartIndex = el.Value.IndexOf(verseStartSearchString);
                        if (verseStartIndex != -1)
                        {
                            int textBreakIndex;
                            int htmlBreakIndex;
                            string textBefore = StringUtils.GetPrevString(el.Value, verseStartIndex + 1, new SearchMissInfo(verseStartIndex, SearchMissInfo.MissMode.CancelOnNextMiss),
                                    out textBreakIndex, out htmlBreakIndex, StringSearchIgnorance.None, StringSearchMode.NotSpecified).Replace("&nbsp;", "");

                            if (textBefore.Length == 0)  // чтоб убедиться, что мы взяли текст в начале строки
                            {
                                int verseEndIndex = el.Value.IndexOf("<", verseStartIndex + 1);

                                if (verseEndIndex != -1)
                                {
                                    string verse = el.Value.Substring(verseStartIndex + verseStartSearchString.Length, verseEndIndex - verseStartIndex - verseStartSearchString.Length);

                                    if (verse == commentText)
                                        return (string)el.Parent.Attribute("objectID");
                                }
                            }
                        }
                    }
                }
            }            

            return null;
        }

        private string GetCommentPageName(string commentLink)
        {
            string result = SettingsManager.Instance.PageName_DefaultComments;
            string beginSearchString = ".one#";
            string endSearchString = ".%20%5b";
            int i = commentLink.IndexOf(beginSearchString);

            if (i != -1)
            {
                int ii = commentLink.IndexOf(endSearchString, i + 1);

                if (ii != -1)
                {
                    result = commentLink.Substring(i + beginSearchString.Length, ii - i - beginSearchString.Length);
                    result = Uri.UnescapeDataString(result);
                }
            }

            return result;
        }

        public void Dispose()
        {
            OneNoteUtils.ReleaseOneNoteApp(ref _oneNoteApp);
            _form = null;
        }
    }
}
