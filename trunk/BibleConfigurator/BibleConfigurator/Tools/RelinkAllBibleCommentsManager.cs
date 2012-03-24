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
    public class RelinkAllBibleCommentsManager
    {
        private Application _oneNoteApp;
        private MainForm _form;

        public RelinkAllBibleCommentsManager(Application oneNoteApp, MainForm form)
        {
            _oneNoteApp = oneNoteApp;
            _form = form;
        }

        public void RelinkAllBibleComments()
        {
            if (!SettingsManager.Instance.IsConfigured(_oneNoteApp))
            {
                Logger.LogError(Constants.Error_SystemIsNotConfigures);
                return;
            }   

            try
            {
                BibleCommon.Services.Logger.Init("RelinkAllBibleCommentsManager");

                _form.PrepareForExternalProcessing(1255, 1, "Старт обновления ссылок на комментарии.");

                //для тестирования
                //RelinkPageComments(_oneNoteApp.Windows.CurrentWindow.CurrentSectionId, 
                //    _oneNoteApp.Windows.CurrentWindow.CurrentPageId, 
                //    OneNoteUtils.GetHierarchyElementName(_oneNoteApp, _oneNoteApp.Windows.CurrentWindow.CurrentPageId));
                //return;

                NotebookIteratorHelper.Iterate(_oneNoteApp,
                    SettingsManager.Instance.NotebookId_Bible, SettingsManager.Instance.SectionGroupId_Bible, pageInfo =>
                        {
                            try
                            {
                                RelinkPageComments(pageInfo.SectionId, pageInfo.Id, pageInfo.Title);
                            }
                            catch (Exception ex)
                            {
                                Logger.LogError(ex.ToString());
                            }

                            if (_form.StopExternalProcess)
                                throw new ProcessAbortedByUserException();
                        });
                
            }
            catch (ProcessAbortedByUserException)
            {
                BibleCommon.Services.Logger.LogMessage("Process aborted by user");
            }
            finally
            {                
                BibleCommon.Services.Logger.Done();

                _form.ExternalProcessingDone("Обновление ссылок на комментарии успешно завершено.");
            }
        }

        private void RelinkPageComments(string sectionId, string pageId, string pageName)
        {
            _form.PerformProgressStep(string.Format("Обработка страницы '{0}'", pageName));

            string pageContent;
            XmlNamespaceManager xnm;            
            _oneNoteApp.GetPageContent(pageId, out pageContent);
            XDocument pageDocument = OneNoteUtils.GetXDocument(pageContent, out xnm);            
            bool wasModified = false;

            foreach (XElement textElement in pageDocument.Root.XPathSelectElements("one:Outline/one:OEChildren/one:OE/one:Table/one:Row/one:Cell[1]/one:OEChildren/one:OE/one:T", xnm))
            {
                OneNoteUtils.NormalizaTextElement(textElement);

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
                OneNoteUtils.UpdatePageContentSafe(_oneNoteApp, pageDocument);
        }

        private bool RelinkPageComment(string bibleSectionId, string biblePageId, string biblePageName, XElement textElement, int linkIndex, int linkEnd)
        {
            string commentLink = textElement.Value.Substring(linkIndex, linkEnd - linkIndex + "</a>".Length);
            string commentText = StringUtils.GetText(commentLink);

            string commentPageName = GetCommentPageName(commentLink);
            string commentPageId = OneNoteProxy.Instance.GetCommentPageId(_oneNoteApp, bibleSectionId, biblePageId, biblePageName, commentPageName);
            string commentObjectId = GetComentObjectId(commentPageId, commentText, null, 0);

            if (!string.IsNullOrEmpty(commentObjectId))
            {
                string newCommentLink = OneNoteUtils.GenerateHref(_oneNoteApp, commentText, commentPageId, commentObjectId);

                textElement.Value = textElement.Value.Replace(commentLink, newCommentLink);               

                return true;
            }

            return false;
        }

        private string GetComentObjectId(string commentPageId, string commentText, string textElementId, int startIndex)
        {            
            OneNoteProxy.PageContent pageDoc = OneNoteProxy.Instance.GetPageContent(_oneNoteApp, commentPageId, OneNoteProxy.PageType.CommentPage);

            foreach (XElement el in pageDoc.Content.Root.XPathSelectElements(string.Format("one:Outline/one:OEChildren/one:OE{0}/one:T", 
                !string.IsNullOrEmpty(textElementId) ? string.Format("[@objectID='{0}']", textElementId) : string.Empty), pageDoc.Xnm))
            {
                OneNoteUtils.NormalizaTextElement(el);

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

                            if (textBefore.Length == 0)  // чотб убедиться, что мы взяли текст в начале строки
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
    }
}
