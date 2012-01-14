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

namespace BibleVerseLinkerEx
{
    public class VerseLinker
    {
        private static readonly string[] PointerStrings = new string[] { "text-decoration:underline", "text-decoration: underline", "text-decoration:\nunderline" };        

        public string DescriptionPageName { get; set; }

        /// <summary>
        /// если false, то не ищем подчёркнутый текст и просто создаём страницу комментариев, которая потом проиндексируется BibleNoteLinker-ом!
        /// </summary>
        public bool SearchForUnderlineText { get; set; }

        private Application _onenoteApp = null;
        public Application OneNoteApp
        {
            get
            {
                return _onenoteApp;
            }
        }

        public VerseLinker()
        {
            _onenoteApp = new Application();
            DescriptionPageName = SettingsManager.Instance.DescriptionPageDefaultName;
            SearchForUnderlineText = true;
        }

        /// <summary>
        /// Возвращает элемент текущей страницы, в котором есть строка, соответствующая PointerFilters
        /// </summary>
        /// <returns></returns>
        private XElement FindPointerElement(string pageId, out XDocument document)
        {
            string pageContentXml;
            OneNoteApp.GetPageContent(pageId, out pageContentXml);

            XmlNamespaceManager xnm;
            document = Utils.GetXDocument(pageContentXml, out xnm);
            XElement pointerElement = null;

            if (SearchForUnderlineText)
            {
                string[] searchPatterns = new string[] { 
                "/one:Page/one:Outline/one:OEChildren/one:OE/one:Table/one:Row/one:Cell/one:OEChildren/one:OE/one:T[contains(.,'{0}')]",
                "/one:Page/one:Outline/one:OEChildren/one:OE/one:T[contains(.,'{0}')]" };

                foreach (string f in PointerStrings)
                {
                    foreach (string pattern in searchPatterns)
                    {
                        pointerElement = document.XPathSelectElement(string.Format(pattern, f), xnm);

                        if (pointerElement != null)
                        {
                            pointerElement.Value = pointerElement.Value.Replace("\n", " ");
                            break;
                        }
                    }

                    if (pointerElement != null)
                        break;
                }
            }

            return pointerElement;
        }

        private XElement GetLastPageObject(string pageId, int? position)
        {
            string pageContentXml;
            OneNoteApp.GetPageContent(pageId, out pageContentXml);

            XmlNamespaceManager xnm;
            XDocument document = Utils.GetXDocument(pageContentXml, out xnm);

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
                XElement pointerElement = FindPointerElement(currentPageId, out currentPageDocument);
                string currentPageName = (string)currentPageDocument.Root.Attribute("name");

                if (!SearchForUnderlineText || pointerElement != null)
                {
                    string pointerValueString = null;
                    string pointerString = null;
                    string currentObjectId = null;
                    int? verseNumber = null;

                    if (pointerElement != null)
                    {
                        pointerString = CutPointerString(pointerElement.Value, out pointerValueString);
                        verseNumber = GetVerseNumber(pointerElement.Value);
                        currentObjectId = (string)pointerElement.Parent.Attribute("objectID");
                    }

                    if (!SearchForUnderlineText || !string.IsNullOrEmpty(pointerString))
                    {
                        string verseLinkPageId = null;
                        try
                        {
                            verseLinkPageId = BibleCommon.Services.VerseLinkManager.FindVerseLinkPageAndCreateIfNeeded(OneNoteApp, currentNotebookId, currentSectionId,
                                currentPageId, currentPageName, DescriptionPageName);                            
                        }
                        catch (Exception ex)
                        {
                            Logger.LogError(ex.Message);
                        }

                        if (!string.IsNullOrEmpty(verseLinkPageId))
                        {
                            string newObjectContent = string.Empty;
                            if (!string.IsNullOrEmpty(pointerString))
                                newObjectContent = GetNewObjectContent(currentPageId, currentObjectId, pointerValueString, verseNumber);

                            string objectId = UpdateDescriptionPage(verseLinkPageId, newObjectContent, verseNumber);

                            if (!string.IsNullOrEmpty(pointerString))
                            {
                                string href = Utils.GenerateHref(OneNoteApp, pointerValueString, verseLinkPageId, objectId);

                                pointerElement.Value = pointerElement.Value.Replace(pointerString, href);
                            }

                            if (SearchForUnderlineText)
                                OneNoteApp.UpdatePageContent(currentPageDocument.ToString());
                            
                            OneNoteApp.NavigateTo(verseLinkPageId, objectId);
                        }
                    }
                    else
                        Logger.LogError("Не удалось выделить подчёркнутый текст");
                }
                else
                    Logger.LogError("Подчёркнутый текст не найден");
            }
            else
                Logger.LogError("Программа OneNote не запущена");
        }

        private int? GetVerseNumber(string textElementValue)
        {
            int? result = null;
            if (textElementValue.StartsWith("<a href"))                         
            {
                string searchPattern = ">";
                int i = textElementValue.IndexOf(searchPattern);
                if (i != -1)
                    result = Utils.GetStringFirstNumber(textElementValue, i + searchPattern.Length);
            }
            else
                result = Utils.GetStringFirstNumber(textElementValue);

            return result;
        }

        private string GetNewObjectContent(string currentPageId, string currentObjectId, string pointerValueString, int? verseNumber)
        {
            string newContent;

            if (verseNumber.HasValue)
            {
                string linkToCurrentObject;
                OneNoteApp.GetHyperlinkToObject(currentPageId, currentObjectId, out linkToCurrentObject);
                newContent = string.Format("<a href=\"{0}\">:{1}</a>&nbsp;&nbsp;<b>{2}</b>", linkToCurrentObject, verseNumber,
                    verseNumber.ToString() != pointerValueString.Trim() ? pointerValueString : string.Empty);
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
        public string UpdateDescriptionPage(string pageId, string pointerValueString, int? verseNumber)
        {
            string pageContentXml;
            XDocument pageDocument;
            XmlNamespaceManager xnm;
            OneNoteApp.GetPageContent(pageId, out pageContentXml);
            pageDocument = Utils.GetXDocument(pageContentXml, out xnm);
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

            if (verseNumber.HasValue)
            {
                string searchPattern = ">:";
                foreach (XElement commentElement in pageDocument.Root.XPathSelectElements("one:Outline/one:OEChildren/one:OE/one:T", xnm))
                {
                    int i = commentElement.Value.IndexOf(searchPattern);
                    if (i != -1)
                    {
                        int? number = Utils.GetStringFirstNumber(commentElement.Value, i + searchPattern.Length);

                        if (number.HasValue)
                        {
                            if (number > verseNumber)
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

            OneNoteApp.UpdatePageContent(pageDocument.ToString());

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
                position.Attribute("y").Value = commentPosition.ToString();

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
                string s = attribute.Value;

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

        /// <summary>
        /// Возвращает html строку - подчёркнутый текст
        /// </summary>
        /// <param name="sourceString">сожержимое всего pointerElement</param>
        /// <param name="pointerValueString">сам текст (не html)</param>
        /// <returns></returns>
        private static string CutPointerString(string sourceString, out string pointerValueString)
        {
            pointerValueString = string.Empty;
            string result = string.Empty;

            int index = sourceString.IndexOf(PointerStrings[0]);
            if (index == -1)
                index = sourceString.IndexOf(PointerStrings[1]);

            if (index != -1)
            {
                string leftPart = sourceString.Substring(0, index);

                int firstLetterIndex = leftPart.LastIndexOf("<");
                int lastLetterIndex = sourceString.IndexOf(">", index);
                if (lastLetterIndex != -1)
                {
                    int i = sourceString.IndexOf("<", lastLetterIndex + 1);
                    if (i != -1)
                        pointerValueString = sourceString.Substring(lastLetterIndex + 1, i - lastLetterIndex - 1);

                    lastLetterIndex = sourceString.IndexOf(">", lastLetterIndex + 1);
                }

                if (firstLetterIndex != -1 && lastLetterIndex != -1)
                {
                    result = sourceString.Substring(firstLetterIndex, lastLetterIndex - firstLetterIndex + 1);

                    string otherString = sourceString.Substring(lastLetterIndex + 1);

                    string otherPointerValue;
                    string otherPointerString = CutPointerString(otherString, out otherPointerValue);

                    if (!string.IsNullOrEmpty(otherPointerString))
                    {
                        result += otherPointerString;
                        pointerValueString += otherPointerValue;
                    }
                }
            }

            return result;
        }
    }
}
