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

namespace BibleVerseLinkerEx
{
    public class VerseLinker: IDisposable
    {
        public string DescriptionPageName { get; set; }

        private Application _oneNoteApp = null;
        public Application OneNoteApp
        {
            get
            {
                return _oneNoteApp;
            }
        }

        public VerseLinker()
        {
            _oneNoteApp = new Application();
            DescriptionPageName = SettingsManager.Instance.PageName_DefaultComments;         
        }

        /// <summary>
        /// Возвращает выделенный текст
        /// </summary>
        /// <returns></returns>
        private XElement FindSelectedText(string pageId, out XDocument document,
            out VerseNumber verseNumber, out string currentObjectId, out XmlNamespaceManager xnm)
        {
            verseNumber = null;
            currentObjectId = null;

            string pageContentXml;
            OneNoteApp.GetPageContent(pageId, out pageContentXml, PageInfo.piSelection, Constants.CurrentOneNoteSchema);

            document = OneNoteUtils.GetXDocument(pageContentXml, out xnm);
            XElement pointerElement = document.Root.XPathSelectElement("//one:T[@selected='all']", xnm);

            if (pointerElement != null)
            {
                OneNoteUtils.NormalizeTextElement(pointerElement);
                verseNumber = VerseNumber.GetFromVerseText(pointerElement.Parent.Value);
                currentObjectId = (string)pointerElement.Parent.Attribute("objectID");                

                return pointerElement;
            }

            return null;
        }

        private XElement GetLastPageObject(string pageId, int? position)
        {
            _oneNoteApp.SyncHierarchy(pageId);

            string pageContentXml;
            OneNoteApp.GetPageContent(pageId, out pageContentXml, PageInfo.piBasic, Constants.CurrentOneNoteSchema);

            XmlNamespaceManager xnm;
            XDocument document = OneNoteUtils.GetXDocument(pageContentXml, out xnm);

            XElement result = null;
            
            if (position.HasValue)
                result = document.Root.XPathSelectElement(string.Format("one:Outline[{0}]", position), xnm);

            if (result == null)
                result = document.Root.XPathSelectElement("one:Title", xnm); 

            return result;
        }

        public void Do()
        {
            if (OneNoteApp.Windows.CurrentWindow != null)
            {
                string currentPageId = OneNoteApp.Windows.CurrentWindow.CurrentPageId;
                string currentSectionId = OneNoteApp.Windows.CurrentWindow.CurrentSectionId;
                string currentNotebookId = OneNoteApp.Windows.CurrentWindow.CurrentNotebookId;

                XDocument currentPageDocument;
                XmlNamespaceManager xnm;
                VerseNumber verseNumber;
                string currentObjectId;
                XElement selectedElement = FindSelectedText(currentPageId, out currentPageDocument, out verseNumber, out currentObjectId, out xnm);
                string selectedHtml = selectedElement != null ? ShellText(selectedElement.Value) : string.Empty;                
                string selectedText = ShellText(StringUtils.GetText(selectedHtml, SettingsManager.Instance.CurrentModule.BibleStructure.Alphabet));
                bool selectedTextFound = !string.IsNullOrEmpty(selectedText);

                if (selectedTextFound)
                {
                    try
                    {
                        BibleCommon.Services.OneNoteLocker.UnlockCurrentSection(OneNoteApp);
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
                    verseLinkPageId = BibleCommon.Services.VerseLinkManager.FindVerseLinkPageAndCreateIfNeeded(OneNoteApp, currentSectionId,
                        currentPageId, currentPageName, DescriptionPageName, false);
                }
                catch (Exception ex)
                {
                    Logger.LogError(ex.Message);
                }

                if (!string.IsNullOrEmpty(verseLinkPageId))
                {
                    string newObjectContent = string.Empty;
                    if (selectedTextFound)
                        newObjectContent = GetNewObjectContent(currentPageId, currentObjectId, selectedText, verseNumber);

                    string objectId = UpdateDescriptionPage(verseLinkPageId, newObjectContent, verseNumber);

                    if (selectedTextFound)
                    {
                        string href = OneNoteUtils.GenerateHref(OneNoteApp, selectedHtml, verseLinkPageId, objectId);

                        string selectedValue = selectedElement.Value;
                        selectedElement.Value = string.Empty;
                        selectedElement.Add(new XCData(selectedValue.Replace(selectedHtml, href)));
                    
                        OneNoteUtils.UpdatePageContentSafe(_oneNoteApp, currentPageDocument, xnm);
                    }

                    OneNoteApp.NavigateTo(verseLinkPageId, objectId);
                }
            }
            else
                Logger.LogError(BibleCommon.Resources.Constants.VerseLinkerOneNoteNotStarted);
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
            foreach (var sortPageInfo in OneNoteProxy.Instance.SortVerseLinkPagesInfo)
            {
                VerseLinkManager.SortVerseLinkPages(_oneNoteApp,
                    sortPageInfo.SectionId, sortPageInfo.PageId, sortPageInfo.ParentPageId, sortPageInfo.PageLevel);
            }

            OneNoteProxy.Instance.CommitAllModifiedHierarchy(_oneNoteApp, null, null);
        }

        private string GetNewObjectContent(string currentPageId, string currentObjectId, string pointerValueString, VerseNumber verseNumber)
        {
            string newContent;

            if (verseNumber != null)
            {
                bool pointerValueIsVerseNumber = verseNumber.ToString() == pointerValueString;
                string linkToCurrentObject;
                OneNoteApp.GetHyperlinkToObject(currentPageId, currentObjectId, out linkToCurrentObject);
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
        /// Возвращает добавленный objectId
        /// </summary>
        /// <param name="pageId"></param>
        /// <param name="pointerValueString"></param>
        /// <returns></returns>
        public string UpdateDescriptionPage(string pageId, string pointerValueString, VerseNumber verseNumber)
        {
            string pageContentXml;
            XDocument pageDocument;
            XmlNamespaceManager xnm;
            OneNoteApp.GetPageContent(pageId, out pageContentXml, PageInfo.piBasic, Constants.CurrentOneNoteSchema);
            pageDocument = OneNoteUtils.GetXDocument(pageContentXml, out xnm);
            XNamespace nms = XNamespace.Get(Constants.OneNoteXmlNs);


            XElement newCommentElement = new XElement(nms + "Outline",
                                            new XElement(nms + "Size", new XAttribute("width", "500"), new XAttribute("height", 15)),
                                            new XElement(nms + "OEChildren",
                                                new XElement(nms + "OE",
                                                    new XElement(nms + "T",
                                                        new XCData(
                                                            pointerValueString
                                                            )))));

            XElement prevComment = null;                    

            if (verseNumber != null)
            {
                string searchPattern = ">:";
                foreach (XElement commentElement in pageDocument.Root.XPathSelectElements("one:Outline/one:OEChildren/one:OE/one:T", xnm))
                {
                    int i = commentElement.Value.IndexOf(searchPattern);
                    if (i != -1)
                    {
                        int? number = StringUtils.GetStringFirstNumber(commentElement.Value, i + searchPattern.Length);

                        if (number.HasValue)
                        {
                            if (number > verseNumber.Verse)
                                break;
                            prevComment = commentElement.Parent.Parent.Parent;
                        }
                    }
                    else
                        break;
                }                                

                if (prevComment == null)
                {
                    prevComment = pageDocument.Root.XPathSelectElement("one:Title", xnm);                    
                    SetPositionYForComment(newCommentElement, 87, xnm, nms);             
                }
                else
                {
                    SetPositionYForComment(newCommentElement, prevComment, xnm, nms);
                }

                prevComment.AddAfterSelf(newCommentElement);

                prevComment = newCommentElement;
                foreach (XElement nextComment in newCommentElement.ElementsAfterSelf())
                {
                    SetPositionYForComment(nextComment, prevComment, xnm, nms);
                    prevComment = nextComment;
                }
            }
            else
                pageDocument.Root.Add(newCommentElement);

            OneNoteUtils.UpdatePageContentSafe(OneNoteApp, pageDocument, xnm);

            XElement addedObject = GetLastPageObject(pageId, GetOutlinePosition(pageDocument, newCommentElement, xnm));

            if (addedObject != null)
            {
                return (string)addedObject.Attribute("objectID");
            }

            return string.Empty;
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
            _oneNoteApp = null;
        }
    }
}
