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

namespace BibleConfigurator.Tools
{
    public class RelinkAllBibleCommentsManager
    {
        public class CommentPageId
        {
            public string BibleSectionId { get; set; }
            public string BiblePageId { get; set; }
            public string BiblePageName { get; set; }
            public string CommentsPageName { get; set; }

            public override int GetHashCode()
            {
                return BibleSectionId.GetHashCode() ^ BiblePageId.GetHashCode() ^ BiblePageName.GetHashCode() ^ CommentsPageName.GetHashCode();
            }

            public override bool Equals(object obj)
            {
                CommentPageId otherObject = (CommentPageId)obj;
                return BibleSectionId == otherObject.BibleSectionId
                    && BiblePageId == otherObject.BiblePageId
                    && BiblePageName == otherObject.BiblePageName
                    && CommentsPageName == otherObject.CommentsPageName;
            }
        }

        private Dictionary<CommentPageId, string> _commentPagesIds = new Dictionary<CommentPageId, string>();

        private Application _oneNoteApp;
        private MainForm _form;

        public RelinkAllBibleCommentsManager(Application oneNoteApp, MainForm form)
        {
            _oneNoteApp = oneNoteApp;
            _form = form;
        }

        public void RelinkAllBibleComments()
        {
            try
            {
                BibleCommon.Services.Logger.Init("RelinkAllBibleCommentsManager");

                _form.PrepareForExternalProcessing(1255, 1, "Старт обновления ссылок на комментарии.");

                ProcessNotebook(SettingsManager.Instance.NotebookId_Bible, SettingsManager.Instance.SectionGroupId_Bible);
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

        private void ProcessNotebook(string notebookId, string sectionGroupId)
        {
            BibleCommon.Services.Logger.LogMessage("Обработка записной книжки: '{0}'", OneNoteUtils.GetHierarchyElementName(_oneNoteApp, notebookId));  // чтобы точно убедиться

            string hierarchyXml;
            _oneNoteApp.GetHierarchy(notebookId, HierarchyScope.hsPages, out hierarchyXml);
            XmlNamespaceManager xnm;
            XDocument notebookDoc = OneNoteUtils.GetXDocument(hierarchyXml, out xnm);

            BibleCommon.Services.Logger.MoveLevel(1);
            ProcessRootSectionGroup(notebookId, notebookDoc, sectionGroupId, xnm);
            BibleCommon.Services.Logger.MoveLevel(-1);
        }

        private void ProcessRootSectionGroup(string notebookId, XDocument doc, string sectionGroupId, XmlNamespaceManager xnm)
        {
            XElement sectionGroup = string.IsNullOrEmpty(sectionGroupId)
                                        ? doc.Root
                                        : doc.Root.XPathSelectElement(
                                                string.Format("one:SectionGroup[@ID='{0}']", sectionGroupId), xnm);

            if (sectionGroup != null)
                ProcessSectionGroup(sectionGroup, sectionGroupId, notebookId, xnm);
            else
                BibleCommon.Services.Logger.LogError("Не удаётся найти группу секций '{0}'", sectionGroupId);
        }

        private  void ProcessSectionGroup(XElement sectionGroup, string sectionGroupId,
            string notebookId, XmlNamespaceManager xnm)
        {
            string sectionGroupName = (string)sectionGroup.Attribute("name");

            if (!string.IsNullOrEmpty(sectionGroupName))
            {
                BibleCommon.Services.Logger.LogMessage("Обработка группы секций '{0}'", sectionGroupName);
                BibleCommon.Services.Logger.MoveLevel(1);
            }

            foreach (var subSectionGroup in sectionGroup.XPathSelectElements("one:SectionGroup", xnm))
            {
                string subSectionGroupName = (string)subSectionGroup.Attribute("name");
                ProcessSectionGroup(subSectionGroup, subSectionGroupName, notebookId, xnm);
            }

            foreach (var subSection in sectionGroup.XPathSelectElements("one:Section", xnm))
            {
                ProcessSection(subSection, sectionGroupId, notebookId, xnm);
            }

            if (!string.IsNullOrEmpty(sectionGroupName))
            {
                BibleCommon.Services.Logger.MoveLevel(-1);
            }
        }

        private void ProcessSection(XElement section, string sectionGroupId,
           string notebookId, XmlNamespaceManager xnm)
        {
            string sectionId = (string)section.Attribute("ID");
            string sectionName = (string)section.Attribute("name");

            BibleCommon.Services.Logger.LogMessage("Обработка секции '{0}'", sectionName);
            BibleCommon.Services.Logger.MoveLevel(1);            

            foreach (var page in section.XPathSelectElements("one:Page", xnm))
            {
                string pageId = (string)page.Attribute("ID");
                string pageName = (string)page.Attribute("name");

                BibleCommon.Services.Logger.LogMessage("Обработка страницы '{0}'", pageName);

                BibleCommon.Services.Logger.MoveLevel(1);

                RelinkPageComments(sectionGroupId, sectionId, pageId, pageName);
                
                BibleCommon.Services.Logger.MoveLevel(-1);

                if (_form.StopExternalProcess)
                    throw new ProcessAbortedByUserException();
            }

            BibleCommon.Services.Logger.MoveLevel(-1);
        }

        private void RelinkPageComments(string sectionGroupId, string sectionId, string pageId, string pageName)
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
                _oneNoteApp.UpdatePageContent(pageDocument.ToString());
        }

        private bool RelinkPageComment(string bibleSectionId, string biblePageId, string biblePageName, XElement textElement, int linkIndex, int linkEnd)
        {
            string commentLink = textElement.Value.Substring(linkIndex, linkEnd - linkIndex + "</a>".Length);
            string commentText = GetLinkText(commentLink);

            string commentPageName = GetCommentPageName(commentLink);
            string commentPageId = GetCommentPageId(bibleSectionId, biblePageId, biblePageName, commentPageName);
            string commentObjectId = GetComentobjectId(commentPageId, commentText);

            if (!string.IsNullOrEmpty(commentObjectId))
            {
                string newCommentLink = OneNoteUtils.GenerateHref(_oneNoteApp, commentText, commentPageId, commentObjectId);

                textElement.Value = textElement.Value.Replace(commentLink, newCommentLink);               

                return true;
            }

            return false;
        }

        private string GetComentobjectId(string commentPageId, string commentText)
        {
            XmlNamespaceManager xnm;
            XDocument pageDoc = OneNoteUtils.GetXDocument(OneNoteProxy.Instance.GetPageContent(_oneNoteApp, commentPageId), out xnm);

            foreach (XElement el in pageDoc.Root.XPathSelectElements("one:Outline/one:OEChildren/one:OE/one:T", xnm))
            {
                OneNoteUtils.NormalizaTextElement(el);

                bool needToSearchVerse = true;
                int boldTagIndex = el.Value.IndexOf("font-weight:bold");
                if (boldTagIndex != -1)
                {
                    boldTagIndex = el.Value.IndexOf(">", boldTagIndex + 1);

                    if (boldTagIndex != -1)
                    {
                        int breakIndex;
                        string textBefore = StringUtils.GetPrevString(el.Value, boldTagIndex + 1, new SearchMissInfo(boldTagIndex, SearchMissInfo.MissMode.CancelOnNextMiss),
                                out breakIndex, StringSearchIgnorance.None, StringSearchMode.NotSpecified).Replace("&nbsp;", "");

                        if (textBefore.Length <= 5)  // чотб убедиться, что мы взяли текст в начале строки
                        {
                            int boldEndIndex = el.Value.IndexOf("</span>", boldTagIndex + 1);

                            if (boldEndIndex != -1)
                            {
                                string commentValue = el.Value.Substring(boldTagIndex + 1, boldEndIndex - boldTagIndex - 1);
                                if (commentValue == commentText)
                                    return (string)el.Parent.Attribute("objectID");
                                else
                                    needToSearchVerse = false;  // это точно не стих, это просто другой комментарий
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
                            int breakIndex;
                            string textBefore = StringUtils.GetPrevString(el.Value, verseStartIndex + 1, new SearchMissInfo(verseStartIndex, SearchMissInfo.MissMode.CancelOnNextMiss),
                                    out breakIndex, StringSearchIgnorance.None, StringSearchMode.NotSpecified).Replace("&nbsp;", "");

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

        private string GetCommentPageId(string bibleSectionId, string biblePageId, string biblePageName, string commentPageName)
        {
            CommentPageId key = new CommentPageId() { 
                BibleSectionId = bibleSectionId, BiblePageId = biblePageId, BiblePageName = biblePageName, CommentsPageName = commentPageName };
            if (!_commentPagesIds.ContainsKey(key))
            {
                string commentPageId = VerseLinkManager.FindVerseLinkPageAndCreateIfNeeded(_oneNoteApp, bibleSectionId, biblePageId, biblePageName, commentPageName);                
                _commentPagesIds.Add(key, commentPageId);
            }

            return _commentPagesIds[key];
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

        private string GetLinkText(string commentLink)
        {            
            int breakIndex;
            string s = StringUtils.GetNextString(commentLink, -1, new SearchMissInfo(commentLink.Length, SearchMissInfo.MissMode.CancelOnNextMiss),
                out breakIndex, StringSearchIgnorance.None, StringSearchMode.NotSpecified);

            return s;
        }
       
    }
}
