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

namespace TestOneNote
{
    class _Program
    {
        private const string PointerString = "text-decoration:underline";

        void Main(string[] args)
        {
            string pageContentXml;
            //string notebookContentXml;
            var onenoteApp = new Application();

            onenoteApp.GetPageContent(onenoteApp.Windows.CurrentWindow.CurrentPageId, out pageContentXml);

            XDocument xd = XDocument.Parse(pageContentXml);
            XmlNamespaceManager xnm = new XmlNamespaceManager(new NameTable());
            xnm.AddNamespace("one", "http://schemas.microsoft.com/office/onenote/2010/onenote");


            XElement targetElement = xd.XPathSelectElement(string.Format(
                "/one:Page/one:Outline/one:OEChildren/one:OE/one:T[contains(.,'{0}')]", PointerString), xnm);

            string newPageId;
            onenoteApp.CreateNewPage(onenoteApp.Windows.CurrentWindow.CurrentSectionId, out newPageId);

            if (targetElement != null)
            {

                string pointerString = CutPointerString(targetElement.Value);

                if (!string.IsNullOrEmpty(pointerString))
                {




                    targetElement.Value = targetElement.Value.Replace(pointerString,
     "<a href=\"onenote:#gds&amp;section-id={FDBC8A3D-69E1-43CF-9273-E248686C366B}&amp;page-id={33CD7A13-5367-413B-852B-B231A946149B}&amp;object-id={D4483700-B4F2-43D0-B999-1A409567543F}&amp;27&amp;base-path=https://azaoj4.docs.live.net/55c10564bd4d0a7b/%5e.Documents/Holy%20Bible/Новый%20раздел%201.one\">sdf</a>");
                    onenoteApp.UpdatePageContent(xd.ToString());
                }
            }


            Console.ReadKey();
        }

        //private static void CreateNewPage()
        //{
        //    onenoteApp.GetHierarchy(onenoteApp.Windows.CurrentWindow.CurrentNotebookId, HierarchyScope.hsSections, out notebookContentXml);
        //}

        private static string CutPointerString(string s)
        {
            //"Qwertyuiop <span\nstyle='font-style:italic;text-decoration:underline'>123</span> aaaaaaaaaaa"

            int index = s.IndexOf(PointerString);
            string leftPart = s.Substring(0, index);

            int firstLetterIndex = leftPart.LastIndexOf("<");
            int lastLetterIndex = s.IndexOf(">", index);
            if (lastLetterIndex != -1)
                lastLetterIndex = s.IndexOf(">", lastLetterIndex + 1);

            if (firstLetterIndex != -1 && lastLetterIndex != -1)
            {
                return s.Substring(firstLetterIndex, lastLetterIndex - firstLetterIndex + 1);
            }

            return string.Empty;
        }
    }
}