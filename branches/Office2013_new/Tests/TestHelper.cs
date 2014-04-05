using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BibleCommon.Services;
using BibleCommon.Helpers;
using System.Xml;
using System.Xml.XPath;
using System.Xml.Linq;
using BibleCommon.Consts;
using BibleCommon.Common;
using Microsoft.Office.Interop.OneNote;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Tests
{
    public static class TestHelper
    {
        public const string TestNotebookName = "Test";
        public const string TestSectionName = "Test";

        public class PageInfo
        {            
            public XDocument PageDoc { get; set; }
            public XmlNamespaceManager Xnm { get; set; }
            public string NotebookId { get; set; }   
        }

        public class AnalyzeResult
        {
            public string OutputHTML { get; set; }
            public List<VersePointer> FoundVerses { get; set; }

        }    

        public static PageInfo GetTestPage(ref Application oneNoteApp, string scenarioName)
        {
            XmlNamespaceManager xnm;

            var notebookId = OneNoteUtils.GetNotebookIdByName(ref oneNoteApp, TestNotebookName, false);
            string testSectionId;
            var testSectionEl = OneNoteUtils.GetHierarchyElementByName(ref oneNoteApp, "Section", TestSectionName, notebookId);
            if (testSectionEl != null)
                testSectionId = (string)testSectionEl.Attribute("ID");
            else
                testSectionId = NotebookGenerator.AddSection(ref oneNoteApp, notebookId, TestSectionName, true);


            var pageDoc = NotebookGenerator.AddPage(ref oneNoteApp, testSectionId, scenarioName, 1, "ru", out xnm);

            return new PageInfo() { PageDoc = pageDoc, Xnm = xnm, NotebookId = notebookId };
        }

        public static void DeleteTestPage(ref Application oneNoteApp, string pageId)
        {
            NotebookGenerator.DeleteHierarchy(ref oneNoteApp, pageId);            
        }        

        public static AnalyzeResult AnalyzeString(ref Application oneNoteApp, string scenarioName, string inputHtml)
        {   
            XmlNamespaceManager xnm;
            var nms = XNamespace.Get(Constants.OneNoteXmlNs);

            var pageInfo = GetTestPage(ref oneNoteApp, scenarioName);                        

            var textEl = new XElement(nms + "Outline",
                            new XElement(nms + "OEChildren",
                                new XElement(nms + "OE",
                                    new XElement(nms + "T",
                                        new XCData(inputHtml)))));

            pageInfo.PageDoc.Root.Add(textEl);
            
            OneNoteUtils.UpdatePageContentSafe(ref oneNoteApp, pageInfo.PageDoc, pageInfo.Xnm);
            var pageId = (string)pageInfo.PageDoc.Root.Attribute("ID");

            var linkManager = new NoteLinkManager();
            linkManager.LinkPageVerses(ref oneNoteApp, pageInfo.NotebookId, pageId, NoteLinkManager.AnalyzeDepth.SetVersesLinks, false, null);
            var i = linkManager.FoundVerses;
            ApplicationCache.Instance.CommitAllModifiedPages(ref oneNoteApp, true, null, null, null);

            var pageDoc = OneNoteUtils.GetPageContent(ref oneNoteApp, pageId, out xnm);
            var tEl = pageDoc.Root.XPathSelectElement("one:Outline/one:OEChildren/one:OE/one:T", xnm);

            DeleteTestPage(ref oneNoteApp, pageId);

            return new AnalyzeResult() { OutputHTML = tEl.Value, FoundVerses = linkManager.FoundVerses };
        }

        public static void CheckVerses(AnalyzeResult result, params string[] verses)
        {
            Assert.IsTrue(result.FoundVerses.Count == verses.Length);
            foreach(var verse in verses)
                Assert.IsTrue(result.FoundVerses.Contains(new VersePointer(verse)));
        }
    }
}
