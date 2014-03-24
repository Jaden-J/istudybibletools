using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.OneNote;
using System.Xml.Linq;
using System.Xml;
using System.Xml.XPath;
using System.Threading;
using BibleCommon;
using BibleCommon.Helpers;
using BibleCommon.Consts;
using BibleCommon.Services;
using System.Runtime.InteropServices;
using BibleCommon.Common;
using BibleCommon.Handlers;
using System.Diagnostics;

namespace BibleVerseLinkerEx
{
    public class VerseLinker: IDisposable
    {
        public string DescriptionPageName { get; set; }

        private Application _oneNoteApp;
        public Application OneNoteApp
        {
            get
            {
                return _oneNoteApp;
            }
        }

        public VerseLinker()
        {
            _oneNoteApp = OneNoteUtils.CreateOneNoteAppSafe();
            DescriptionPageName = SettingsManager.Instance.PageName_DefaultComments;         
        }

        /// <summary>
        /// Возвращает выделенный текст
        /// </summary>
        /// <returns></returns>
        private XElement FindSelectedText(string pageId, out XDocument document,
            out VerseNumber? verseNumber, out VersePointer versePointer, out string currentObjectId, out XmlNamespaceManager xnm)
        {
            verseNumber = null;
            currentObjectId = null;
            versePointer = null;

            string pageContentXml = null;
            OneNoteUtils.UseOneNoteAPI(ref _oneNoteApp, () =>
            {
                _oneNoteApp.GetPageContent(pageId, out pageContentXml, PageInfo.piSelection, Constants.CurrentOneNoteSchema);
            });

            document = OneNoteUtils.GetXDocument(pageContentXml, out xnm);
            XElement pointerElement = document.Root.XPathSelectElement("//one:Outline//one:T[@selected=\"all\"]", xnm);

            if (pointerElement != null)
            {
                OneNoteUtils.NormalizeTextElement(pointerElement);
                verseNumber = VerseNumber.GetFromVerseText(pointerElement.Parent.Value);
                currentObjectId = (string)pointerElement.Parent.Attribute("objectID");

                if (verseNumber.HasValue)
                    versePointer = TryToExtractVersePointer(document, verseNumber.Value);             

                return pointerElement;
            }

            return null;
        }

        private VersePointer TryToExtractVersePointer(XDocument document, VerseNumber verseNumber)
        {
            try
            {
                var chapterName = (string)document.Root.Attribute("name");
                var parts = chapterName.Split(new string[] { ". " }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length == 2)                
                    chapterName = string.Format("{0} {1}", parts[1], StringUtils.GetStringFirstNumber(parts[0]));

                var versePointer = new VersePointer(string.Format("{0}:{1}", chapterName, verseNumber));
                if (versePointer.IsValid)
                    return versePointer;
            }
            catch (Exception ex)
            {
                FormLogger.LogError(ex);
            }

            return null;
        }     

        public void Do()
        {
            if (_oneNoteApp.Windows.CurrentWindow != null)
            {
                string currentPageId = _oneNoteApp.Windows.CurrentWindow.CurrentPageId;
                string currentSectionId = _oneNoteApp.Windows.CurrentWindow.CurrentSectionId;
                string currentNotebookId = _oneNoteApp.Windows.CurrentWindow.CurrentNotebookId;                

                XDocument currentPageDocument;
                XmlNamespaceManager xnm;
                VerseNumber? verseNumber;
                VersePointer versePointer;
                string currentObjectId;
                XElement selectedElement = FindSelectedText(currentPageId, out currentPageDocument, out verseNumber, out versePointer, out currentObjectId, out xnm);
                string selectedHtml = selectedElement != null ? ShellText(selectedElement.Value) : string.Empty;                
                string selectedText = ShellText(StringUtils.GetText(selectedHtml));
                bool selectedTextFound = !string.IsNullOrEmpty(selectedText);

                if (selectedTextFound)
                {
                    try
                    {
                        BibleCommon.Services.OneNoteLocker.UnlockCurrentSection(ref _oneNoteApp);                                             
                    }
                    catch (NotSupportedException)
                    {
                        //todo: log it
                    }
                }

                string currentPageName = (string)currentPageDocument.Root.Attribute("name");

                string verseLinkPageId = null;
                try
                {
                    bool pageWasCreated;
                    verseLinkPageId = BibleCommon.Services.VerseLinkManager.FindVerseLinkPageAndCreateIfNeeded(ref _oneNoteApp, currentSectionId,
                        currentPageId, currentPageName, DescriptionPageName, false, out pageWasCreated);
                }
                catch (Exception ex)
                {
                    FormLogger.LogError(OneNoteUtils.ParseErrorAndMakeItMoreUserFriendly(ex.Message));
                }

                if (!string.IsNullOrEmpty(verseLinkPageId))
                {
                    string newObjectContent = string.Empty;
                    if (selectedTextFound)
                        newObjectContent = GetNewObjectContent(currentPageId, currentObjectId, versePointer, selectedText, verseNumber);

                    string bnPId, bnOeId;
                    UpdateDescriptionPage(verseLinkPageId, newObjectContent, verseNumber, out bnPId, out bnOeId);

                    var notebookName = OneNoteUtils.GetHierarchyElementName(ref _oneNoteApp, SettingsManager.Instance.NotebookId_BibleComments);
                    var href = ApplicationCache.Instance.GenerateHref(ref _oneNoteApp, 
                                        new LinkId(notebookName, bnPId, bnOeId) { IdType = IdType.Custom },
                                        new LinkProxyInfo(true, true));                                        

                    if (selectedTextFound)
                    {
                        var link = OneNoteUtils.GetLink(selectedHtml, href);
                        var selectedValue = selectedElement.Value;
                        selectedElement.Value = string.Empty;
                        selectedElement.Add(new XCData(selectedValue.Replace(selectedHtml, link)));

                        OneNoteUtils.UpdatePageContentSafe(ref _oneNoteApp, currentPageDocument, xnm);                        
                    }

                    Process.Start(href);
                }
            }
            else
                FormLogger.LogError(BibleCommon.Resources.Constants.VerseLinkerOneNoteNotStarted);
        }

        private string ShellText(string text)
        {
            if (string.IsNullOrEmpty(text))
                return text;

            return text.Trim(new char[] { ' ', '.', ';', ',', ':' });
        }

        public void SortCommentsPages()
        {
            //Сортировка страниц 'Сводные заметок'
            foreach (var sortPageInfo in ApplicationCache.Instance.SortVerseLinkPagesInfo)
            {
                VerseLinkManager.SortVerseLinkPages(ref _oneNoteApp,
                    sortPageInfo.SectionId, sortPageInfo.PageId, sortPageInfo.ParentPageId, sortPageInfo.PageLevel);
            }

            ApplicationCache.Instance.CommitAllModifiedHierarchy(ref _oneNoteApp, null, null);
        }

        private string GetNewObjectContent(string currentPageId, string currentObjectId, VersePointer versePointer, string pointerValueString, VerseNumber? verseNumber)
        {
            string newContent;

            if (verseNumber != null)
            {   
                var useBibleVerseHandler = SettingsManager.Instance.UseProxyLinksForBibleVerses && versePointer != null;
                string linkToCurrentObject = OneNoteUtils.GetOrGenerateLinkHref(ref _oneNoteApp,
                                                useBibleVerseHandler
                                                    ? OpenBibleVerseHandler.GetCommandUrlStatic(versePointer, SettingsManager.Instance.ModuleShortName) 
                                                    : null,
                                                new LinkId(currentPageId, currentObjectId), new LinkProxyInfo(true, false),
                                                useBibleVerseHandler
                                                    ? null
                                                    : Constants.QueryParameter_BibleVerse);

                bool pointerValueIsVerseNumber = verseNumber.ToString() == pointerValueString;
                newContent = string.Format("<a href=\"{0}\">:{1}</a>{2}<b>{3}</b>", linkToCurrentObject, verseNumber,
                    !pointerValueIsVerseNumber ? "&nbsp" : string.Empty,
                    !pointerValueIsVerseNumber ? pointerValueString : string.Empty);
            }
            else
                newContent = string.Format("<b>{0}</b>", pointerValueString);

            newContent += " - ";

            return newContent;
        }

        /// <summary>
        /// Возвращает добавленный bnOEID
        /// </summary>
        /// <param name="pageId"></param>
        /// <param name="pointerValueString"></param>
        /// <returns></returns>
        public void UpdateDescriptionPage(string pageId, string pointerValueString, VerseNumber? verseNumber, out string bnPId, out string bnOeId)
        {
            XNamespace nms = XNamespace.Get(Constants.OneNoteXmlNs);
            var pageContent = ApplicationCache.Instance.GetPageContent(ref _oneNoteApp, pageId, ApplicationCache.PageType.CommentPage);

            XElement newCommentElement = new XElement(nms + "Outline",
                                            new XElement(nms + "Size", new XAttribute("width", "500"), new XAttribute("height", 15)),
                                            new XElement(nms + "OEChildren",
                                                new XElement(nms + "OE",
                                                    new XElement(nms + "T",
                                                        new XCData(
                                                            pointerValueString
                                                            )))));

            XElement prevComment = null;                    

            if (verseNumber.HasValue)
            {
                string searchPattern = ">:";
                foreach (XElement commentElement in pageContent.Content.Root.XPathSelectElements("one:Outline/one:OEChildren/one:OE/one:T", pageContent.Xnm))
                {
                    int i = commentElement.Value.IndexOf(searchPattern);
                    if (i != -1)
                    {
                        int? number = StringUtils.GetStringFirstNumber(commentElement.Value, i + searchPattern.Length);

                        if (number.HasValue)
                        {
                            if (number > verseNumber.Value.Verse)
                                break;
                            prevComment = commentElement.Parent.Parent.Parent;
                        }
                    }
                    else
                        break;
                }                                

                if (prevComment == null)
                {
                    prevComment = pageContent.Content.Root.XPathSelectElement("one:Title", pageContent.Xnm);
                    SetPositionYForComment(newCommentElement, 87, pageContent.Xnm, nms);             
                }
                else
                {
                    SetPositionYForComment(newCommentElement, prevComment, pageContent.Xnm, nms);
                }

                prevComment.AddAfterSelf(newCommentElement);

                prevComment = newCommentElement;
                foreach (XElement nextComment in newCommentElement.ElementsAfterSelf())
                {
                    SetPositionYForComment(nextComment, prevComment, pageContent.Xnm, nms);
                    prevComment = nextComment;
                }
            }
            else
                pageContent.Content.Root.Add(newCommentElement);

            pageContent.WasModified = true;            
            bnPId = OneNoteProxyLinksHandler.GetOrUpdateBnPId(pageContent.Content.Root, pageContent.Xnm).Id;
            bnOeId = OneNoteProxyLinksHandler.GetOrUpdateBnOeId(
                                newCommentElement.XPathSelectElement("one:OEChildren/one:OE", pageContent.Xnm),
                                pageContent.Xnm).Id;                

            ApplicationCache.Instance.CommitAllModifiedPages(ref _oneNoteApp, true, pc => pc.PageType == ApplicationCache.PageType.CommentPage, null, null);            

            //var addedObject = GetLastPageObject(pageId, GetOutlinePosition(pageContent.Content, newCommentElement, pageContent.Xnm));

            //if (addedObject != null)
            //{
            //    bnOeId = OneNoteProxyLinksHandler.GetOrUpdateBnOeId(addedObject, pageContent.Xnm).Id;                
            //}
            //else
            //    bnOeId = string.Empty;
        }

        private XElement GetLastPageObject(string pageId, int? position)
        {
            OneNoteUtils.UseOneNoteAPI(ref _oneNoteApp, () =>
            {
                _oneNoteApp.SyncHierarchy(pageId);
            });

            XElement result = null;

            var pageInfo = ApplicationCache.Instance.GetPageContent(ref _oneNoteApp, pageId, ApplicationCache.PageType.CommentPage, true, PageInfo.piBasic, false);            

            if (position.HasValue)
                result = pageInfo.Content.Root.XPathSelectElement(string.Format("one:Outline[{0}]", position), pageInfo.Xnm);

            if (result == null)
                result = pageInfo.Content.Root.XPathSelectElement("one:Title", pageInfo.Xnm);

            return result;
        }

        private static int GetOutlinePosition(XDocument document, XElement outline, XmlNamespaceManager xnm)
        {
            int i = 0;
            foreach (XElement el in document.Root.XPathSelectElements("one:Outline", xnm))
            {
                i++;
                if (el == outline)
                    break;
            }

            return i;
        }

        private static int SetPositionYForComment(XElement commentElement, int commentPosition, XmlNamespaceManager xnm, XNamespace nms)
        {
            XElement position = commentElement.XPathSelectElement("one:Position", xnm);

            if (position == null)
                commentElement.AddFirst(new XElement(nms + "Position", new XAttribute("x", 36), new XAttribute("y", commentPosition), new XAttribute("z", 0)));
            else
                position.SetAttributeValue("y", commentPosition);

            return commentPosition;
        }

        private static int SetPositionYForComment(XElement commentElement, XElement prevCommentElement, XmlNamespaceManager xnm, XNamespace nms)
        {
            XAttribute prevPositionY = prevCommentElement.XPathSelectElement("one:Position", xnm).Attribute("y");
            XAttribute prevPositionHeight = prevCommentElement.XPathSelectElement("one:Size", xnm).Attribute("height");

            int commentPosition = ParseIntAttributeValue(prevPositionY).GetValueOrDefault(0) + ParseIntAttributeValue(prevPositionHeight).GetValueOrDefault(0) + 22;

            return SetPositionYForComment(commentElement, commentPosition, xnm, nms);            
        }

        private static int? ParseIntAttributeValue(XAttribute attribute)
        {
            if (attribute != null)
            {
                string s = (string)attribute;

                if (!string.IsNullOrEmpty(s))
                {
                    int i = s.IndexOfAny(new char[] { '.', ',' });
                    if (i != -1)
                        s = s.Substring(0, i);

                    return int.Parse(s);
                }
            }

            return null;
        }        

        public void Dispose()
        {
            OneNoteUtils.ReleaseOneNoteApp(ref _oneNoteApp);
        }
    }
}
