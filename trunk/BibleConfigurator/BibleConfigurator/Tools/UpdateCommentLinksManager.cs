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

        public RelinkAllBibleCommentsManager(Application oneNoteApp)
        {
            _oneNoteApp = oneNoteApp; 
        }

        public void RelinkAllBibleComments()
        {
            BibleCommon.Services.Logger.Init("RelinkAllBibleCommentsManager");

            ProcessNotebook(SettingsManager.Instance.NotebookId_Bible, SettingsManager.Instance.SectionGroupId_Bible);

            BibleCommon.Services.Logger.Done();
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
            }

            BibleCommon.Services.Logger.MoveLevel(-1);
        }

        private void RelinkPageComments(string sectionGroupId, string sectionId, string pageId, string pageName)
        {
            string pageContent;
            XmlNamespaceManager xnm;            
            _oneNoteApp.GetPageContent(pageId, out pageContent);
            XDocument pageDocument = OneNoteUtils.GetXDocument(pageContent, out xnm);            
            bool wasModified = false;

            foreach (XElement rowElement in pageDocument.Root.XPathSelectElements("one:Outline/one:OEChildren/one:OE/one:Table/one:Row/one:Cell[1]", xnm))
            {
                string rowElementValue = rowElement.Value.Replace("\n", " ");                

                int linkIndex = rowElementValue.IndexOf("<a ");

                while (linkIndex > -1)
                {
                    int linkEnd = rowElementValue.IndexOf("</a>", linkIndex + 1);

                    if (linkEnd != -1)
                    {
                        if (RelinkPageComment(sectionId, pageId, pageName, rowElement, rowElementValue, linkIndex, linkEnd))
                            wasModified = true;

                        if (wasModified)
                            break;
                    }

                    linkIndex = rowElementValue.IndexOf("<a ", linkIndex + 1);
                }
                if (wasModified)
                    break;
            }

            if (wasModified)
                _oneNoteApp.UpdatePageContent(pageDocument.ToString());

            //ProgressBar.Progress()
        }

        private bool RelinkPageComment(string bibleSectionId, string biblePageId, string biblePageName, XElement rowElement, string rowElementValue, int linkIndex, int linkEnd)
        {
            string commentLink = rowElementValue.Substring(linkIndex, linkEnd - linkIndex + "</a>".Length);
            string commentText = GetLinkText(commentLink);

            string commentPageName = GetCommentPageName(commentLink);
            string commentPageId = GetCommentPageId(bibleSectionId, biblePageId, biblePageName, commentPageName);
            string commentObjectId = GetComentobjectId(commentPageId, commentText);

            if (!string.IsNullOrEmpty(commentObjectId))
            {
                string newCommentLink = OneNoteUtils.GenerateHref(_oneNoteApp, commentText, commentPageId, commentObjectId);

                rowElement.Value = rowElementValue.Replace(commentLink, newCommentLink);
                    //string.Concat(
                    //            rowElement.Value.Substring(0, linkIndex),
                    //            newCommentLink,
                    //            rowElement.Value.Substring(linkEnd + "</a>".Length));

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
                        string verseStartSearchString = "\">:";
                        int verseStartIndex = el.Value.IndexOf(verseStartSearchString);
                        if (verseStartIndex != -1)
                        {
                            int breakIndex;
                            string textBefore = StringUtils.GetPrevString(el.Value, verseStartIndex + 2, new SearchMissInfo(verseStartIndex, SearchMissInfo.MissMode.CancelOnNextMiss),
                                    out breakIndex, StringSearchIgnorance.None, StringSearchMode.NotSpecified).Replace("&nbsp;", "");

                            if (textBefore.Length == 0)  // чотб убедиться, что мы взяли текст в начале строки
                            {
                                int verseEndIndex = el.Value.IndexOf("</a>", verseStartIndex + 1);

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
