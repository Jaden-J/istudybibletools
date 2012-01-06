using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.OneNote;
using System.Xml.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Xml.XPath;
using System.Xml;


namespace BibleVerseLinker
{
    public class VerseLinkManager
    {
        private static readonly string[] PointerStrings = new string[] { "text-decoration:underline", "text-decoration:\nunderline" };


        private const string OneNoteXmlNs = "http://schemas.microsoft.com/office/onenote/2010/onenote";
        private const string OverviewPageName = "Общий обзор";

        private string _descriptionPageName;

        public string DescriptionPageName
        {
            get
            {
                return _descriptionPageName;
            }
            set
            {
                _descriptionPageName = value;
            }
        }

        private Application _onenoteApp = null;
        public Application OneNoteApp
        {
            get
            {
                return _onenoteApp;
            }
        }

        public VerseLinkManager()
        {
            _onenoteApp = new Application();
            _descriptionPageName = "Комментарии";
        } 

        /// <summary>
        /// Возвращает элемент текущей страницы, в котором есть строка, соответствующая PointerFilters
        /// </summary>
        /// <returns></returns>
        public XElement FindPointerElement(string pageId, out XDocument document)
        {
            string pageContentXml;
            OneNoteApp.GetPageContent(pageId, out pageContentXml);

            XmlNamespaceManager xnm;
            document = GetXDocument(pageContentXml, out xnm);
            XElement pointerElement = null;

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
                        pointerElement.Value = pointerElement.Value.Replace("\n", string.Empty);
                        break;
                    }
                }

                if (pointerElement != null)
                    break;
            }

            return pointerElement;
        }

        private XElement GetLastPageObject(string pageId)
        {
            string pageContentXml;
            OneNoteApp.GetPageContent(pageId, out pageContentXml);

            XmlNamespaceManager xnm;
            XDocument document = GetXDocument(pageContentXml, out xnm);
            XElement el = document.XPathSelectElement("/one:Page/one:Outline[last()]", xnm);

            return el;
        }

        private XDocument GetXDocument(string xml, out XmlNamespaceManager xnm)
        {
            XDocument xd = XDocument.Parse(xml);
            xnm = new XmlNamespaceManager(new NameTable());
            xnm.AddNamespace("one", OneNoteXmlNs);
            return xd;
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

                if (pointerElement != null)
                {
                    string pointerValueString;
                    string pointerString = CutPointerString(pointerElement.Value, out pointerValueString);
                    int? verseNumber = GetStringFirstNumber(pointerElement.Value);
                    string currentObjectId = (string)pointerElement.Parent.Attribute("objectID");

                    if (!string.IsNullOrEmpty(pointerString))
                    {
                        string sectionGroupId = FindDescriptionSectionGroupForCurrentPage(currentNotebookId, currentSectionId);
                        if (!string.IsNullOrEmpty(sectionGroupId))
                        {
                            string sectionId = FindDescriptionSectionForCurrentPage(currentPageName, sectionGroupId);
                            if (!string.IsNullOrEmpty(sectionId))
                            {
                                string pageId = FindDescriptionPageForCurrentPage(sectionId, currentPageId, currentPageName);
                                if (!string.IsNullOrEmpty(pageId))
                                {
                                    string newObjectContent = GetNewObjectContent(currentPageId, currentObjectId, pointerValueString, verseNumber);
                                    string objectId = UpdateDescriptionPage(pageId, newObjectContent);
                                    string href = GenerateHref(pointerValueString, pageId, objectId);

                                    pointerElement.Value = pointerElement.Value.Replace(pointerString, href);
                                    OneNoteApp.UpdatePageContent(currentPageDocument.ToString());
                                    OneNoteApp.NavigateTo(pageId, objectId);
                                }
                                else
                                    Logger.LogError("Не найдена страница для комментариев");
                            }
                            else
                                Logger.LogError("Не найдена секция для комментариев");
                        }
                        else
                            Logger.LogError("Не найдена группа секций для комментариев");
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
        /// возвращает номер, находящийся в начале строки: например вернёт 12 для строки "12 глава"
        /// </summary>
        /// <param name="pointerElement"></param>
        /// <returns></returns>
        private static int? GetStringFirstNumber(string s)
        {
            int i = s.IndexOfAny(new char[] { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9' });
            if (i != -1)
            {
                string d1 = s[i].ToString();
                string d2 = string.Empty;
                string d3 = string.Empty;

                d2 = GetDigit(s, i + 1);
                if (!string.IsNullOrEmpty(d2))
                    d3 = GetDigit(s, i + 2);

                return int.Parse(d1 + d2 + d3);
            }

            return null;
        }

        private static string GetDigit(string s, int index)
        {
            int d;
            if (int.TryParse(s[index].ToString(), out d))
                return d.ToString();

            return string.Empty;
        }

        private string FindDescriptionSectionForCurrentPage(string currentPageName, string targetSectionGroupId)
        {
            XmlNamespaceManager xnm;

            string sectionGroupXml;
            OneNoteApp.GetHierarchy(targetSectionGroupId, HierarchyScope.hsSections, out sectionGroupXml);
            XDocument sectionGroupDocument = GetXDocument(sectionGroupXml, out xnm);
            string sectionGroupName = (string)sectionGroupDocument.Root.Attribute("name");

            if (sectionGroupName.IndexOf(currentPageName) != -1)
                currentPageName = OverviewPageName;

            XElement targetSection = sectionGroupDocument.Root.XPathSelectElement(
                string.Format("one:Section[@name='{0}']", currentPageName), xnm);

            if (targetSection == null)
            {
                CreateDescriptionSectionForCurrentPage(sectionGroupDocument, currentPageName);

                return FindDescriptionSectionForCurrentPage(currentPageName, targetSectionGroupId);  // надо обновить XML
            }
            else
                return (string)targetSection.Attribute("ID");
        }

        private void CreateDescriptionSectionForCurrentPage(XDocument sectionGroupDocument, string pageName)
        {
            XNamespace nms = XNamespace.Get(OneNoteXmlNs);
            XElement targetSection = new XElement(nms + "Section",
                                    new XAttribute("name", pageName));

            if (pageName == OverviewPageName || sectionGroupDocument.Root.Nodes().Count() == 0)
                sectionGroupDocument.Root.AddFirst(targetSection);
            else
            {
                int? pageNameIndex = GetStringFirstNumber(pageName);
                bool wasAdded = false;
                foreach (XElement section in sectionGroupDocument.Root.Nodes())
                {
                    string name = (string)section.Attribute("name");
                    int? otherPageIndex = GetStringFirstNumber(name);

                    if (pageNameIndex.GetValueOrDefault(0) < otherPageIndex.GetValueOrDefault(0))
                    {
                        section.AddBeforeSelf(targetSection);
                        wasAdded = true;
                        break;
                    }
                }

                if (!wasAdded)
                    sectionGroupDocument.Root.Add(targetSection);
            }

            OneNoteApp.UpdateHierarchy(sectionGroupDocument.ToString());
        }


        private string GenerateHref(string title, string pageId, string objectId)
        {
            string link;
            OneNoteApp.GetHyperlinkToObject(pageId, objectId, out link);

            return string.Format("<a href=\"{0}\">{1}</a>", link, title);
        }

        public string FindDescriptionSectionGroupForCurrentPage(string currentNotebookId, string currentSectionId)
        {
            XmlNamespaceManager xnm;

            string notebookContentXml;
            OneNoteApp.GetHierarchy(currentNotebookId, HierarchyScope.hsSections, out notebookContentXml);

            XDocument document = GetXDocument(notebookContentXml, out xnm);

            XElement currentSection = document.XPathSelectElement(string.Format("/one:Notebook/one:SectionGroup/one:SectionGroup/one:Section[@ID='{0}']",
                currentSectionId), xnm);

            if (currentSection != null && currentSection.Parent != null && currentSection.Parent.Parent != null)
            {
                string sectionName = (string)currentSection.Attribute("name");                          // 01. От Матфея
                string sectionGroupName = (string)currentSection.Parent.Attribute("name");              // Новый Завет
                string topSectonGroupName = (string)currentSection.Parent.Parent.Attribute("name");     // Бибия

                XElement targetParentSectionGroup = document.XPathSelectElement(
                    string.Format("/one:Notebook/one:SectionGroup[@name!='{0}']/one:SectionGroup[@name='{1}']",
                    topSectonGroupName, sectionGroupName), xnm);                                        // Изучение Библии/Новый Завет

                if (targetParentSectionGroup != null)
                {
                    XElement targetSectionGroup = targetParentSectionGroup.XPathSelectElement(
                        string.Format("one:SectionGroup[@name='{0}']", sectionName), xnm);             // Изучение Библии/Новый Завет/01. От Матфея

                    if (targetSectionGroup == null)
                    {
                        CreateDescriptionSectionGroupForCurrentPage(document, targetParentSectionGroup, sectionName);

                        return FindDescriptionSectionGroupForCurrentPage(currentNotebookId, currentSectionId);  // надо обновить XML
                    }
                    else
                        return (string)targetSectionGroup.Attribute("ID");
                }

            }

            return string.Empty;
        }

        private void CreateDescriptionSectionGroupForCurrentPage(XDocument document, XElement targetParentSectionGroup, string sectionName)
        {
            XNamespace nms = XNamespace.Get(OneNoteXmlNs);
            XElement targetSectionGroup = new XElement(nms + "SectionGroup",
                                    new XAttribute("name", sectionName));

            targetParentSectionGroup.Add(targetSectionGroup);

            OneNoteApp.UpdateHierarchy(document.ToString());
        }

        public string FindDescriptionPageForCurrentPage(string sectionId, string currentPageId, string currentPageName)
        {
            XmlNamespaceManager xnm;
            string sectionContentXml;
            OneNoteApp.GetHierarchy(sectionId, HierarchyScope.hsPages, out sectionContentXml);
            XDocument sectionDocument = GetXDocument(sectionContentXml, out xnm);

            string pageDisplayName = string.Format("{0}. {1}", DescriptionPageName, currentPageName);

            XElement page = sectionDocument.Root.XPathSelectElement(string.Format("one:Page[@name='{0}']", pageDisplayName), xnm);

            string pageId;

            if (page == null)
            {
                OneNoteApp.CreateNewPage(sectionId, out pageId, NewPageStyle.npsBlankPageWithTitle);

                string linkToCurrentPage;
                OneNoteApp.GetHyperlinkToObject(currentPageId, null, out linkToCurrentPage);
                string pageName = string.Format("{0}. <a style='font-size:10pt;' href='{1}'>{2}</a>", DescriptionPageName, linkToCurrentPage, currentPageName);
                SetPageName(pageId, pageName);
            }
            else
                pageId = (string)page.Attribute("ID");

            return pageId;
        }

        public void SetPageName(string pageId, string pageName)
        {
            XNamespace nms = XNamespace.Get(OneNoteXmlNs);
            XDocument pageDocument = new XDocument(new XElement(nms + "Page",
                            new XAttribute("ID", pageId),
                            new XElement(nms + "Title",
                                new XElement(nms + "OE",
                                    new XElement(nms + "T",
                                        new XCData(
                                            pageName
                                            ))))));

            OneNoteApp.UpdatePageContent(pageDocument.ToString());
        }

        /// <summary>
        /// Возвращает добавленный objectId
        /// </summary>
        /// <param name="pageId"></param>
        /// <param name="pointerValueString"></param>
        /// <returns></returns>
        public string UpdateDescriptionPage(string pageId, string pointerValueString)
        {
            XNamespace nms = XNamespace.Get(OneNoteXmlNs);
            var page = new XDocument(new XElement(nms + "Page",
                                new XAttribute("ID", pageId),
                                new XElement(nms + "Outline",
                                  new XElement(nms + "OEChildren",
                                    new XElement(nms + "OE",
                                      new XElement(nms + "T",
                                        new XCData(
                                            pointerValueString
                                            )))))));

            OneNoteApp.UpdatePageContent(page.ToString());


            XElement addedObject = GetLastPageObject(pageId);

            if (addedObject != null)
            {
                return (string)addedObject.Attribute("objectID");
            }

            return string.Empty;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sourceString"></param>
        /// <param name="pointerValueString"></param>
        /// <returns></returns>
        private static string CutPointerString(string sourceString, out string pointerValueString)
        {
            pointerValueString = string.Empty;
            string result = string.Empty;

            int index = sourceString.IndexOf(PointerStrings[0]);
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
