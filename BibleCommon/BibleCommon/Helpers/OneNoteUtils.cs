using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Xml;
using Microsoft.Office.Interop.OneNote;
using System.Reflection;
using System.IO;
using BibleCommon.Consts;
using BibleCommon.Services;
using System.Xml.XPath;

namespace BibleCommon.Helpers
{
    public static class OneNoteUtils
    {
        public static string GetNotebookIdByName(Application oneNoteApp, string notebookName)
        {
            string xml;
            XmlNamespaceManager xnm;
            oneNoteApp.GetHierarchy(null, HierarchyScope.hsNotebooks, out xml);
            XDocument doc = OneNoteUtils.GetXDocument(xml, out xnm);
            XElement bibleNotebook = doc.Root.XPathSelectElement(string.Format("one:Notebook[@nickname='{0}']", notebookName), xnm);
            if (bibleNotebook == null)
                bibleNotebook = doc.Root.XPathSelectElement(string.Format("one:Notebook[@name='{0}']", notebookName), xnm);
            if (bibleNotebook != null)
            {
                return (string)bibleNotebook.Attribute("ID");
            }

            return string.Empty;
        }     

        public static string GetHierarchyElementName(Application oneNoteApp, string elementId)
        {
            XmlNamespaceManager xnm;
            string xml = OneNoteProxy.Instance.GetHierarchy(oneNoteApp, elementId, HierarchyScope.hsSelf);
            XDocument doc = GetXDocument(xml, out xnm);
            return (string)doc.Root.Attribute("name");
        }

        public static XDocument GetXDocument(string xml, out XmlNamespaceManager xnm)
        {
            XDocument xd = XDocument.Parse(xml);
            xnm = new XmlNamespaceManager(new NameTable());
            xnm.AddNamespace("one", Constants.OneNoteXmlNs);
            return xd;
        }

        // возвращает количество родительских узлов
        public static int GetDepthLevel(XElement element)
        {
            int result = 0;

            if (element.Parent != null)
            {                
                result += 1 + GetDepthLevel(element.Parent);
            }

            return result;
        }      

        public static string GenerateHref(Application oneNoteApp, string title, string pageId, string objectId)        
        {
            string link;
            oneNoteApp.GetHyperlinkToObject(pageId, objectId, out link);

            return string.Format("<a href=\"{0}\">{1}</a>", link, title);
        }

        public static void NormalizaTextElement(XElement textElement)  // must be one:T element
        {
            textElement.Value = textElement.Value.Replace("\n", " ");
        }
    }
}
