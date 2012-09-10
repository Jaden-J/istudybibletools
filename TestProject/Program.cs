using System;
using System.Collections.Generic;
using System.Linq;
using BibleCommon.Common;
using BibleConfigurator.ModuleConverter;
using System.Text;
using BibleCommon.Services;
using BibleCommon.Helpers;
using System.IO;
using System.Diagnostics;
using System.Xml.XPath;
using System.Threading;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Xml;
using System.Xml.Linq;
using BibleCommon.Consts;
using Microsoft.Office.Interop.OneNote;


namespace TestProject
{
    class Program
    {
        private static Microsoft.Office.Interop.OneNote.Application _oneNoteApp;
        private static Microsoft.Office.Interop.OneNote.Application OneNoteApp
        {
            get
            {
                if (_oneNoteApp == null)
                    _oneNoteApp = new Microsoft.Office.Interop.OneNote.Application();

                return _oneNoteApp;
            }
        }

        static void Main(string[] args)
        {
           
            try
            {
                //var result = BibleParallelTranslationConnectorManager.GetParallelBibleInfo("rst", "kjv");
                //int i = result.Count;  
                
                //SearchForEnText();

                //ChangeCurrentPageLocale("ro");

                //TryToUpdateInkNodes();

                ConvertRussianModule();

                //ConvertEnglishModule();

                //ConvertRomanModule();

                //GenerateSummaryOfNotesNotebook();

                //DefaultRusModuleGenerator.GenerateModuleInfo("g:\\manifest.xml", true);

                //GenerateParallelBible();   

                //CreateSupplementalBible();

                //GenerateBookDifferencesFile();

            }
            catch (Exception ex)
            {
                Logger.LogError(ex.ToString());
            }

            Console.ReadKey();
        }

        //private static void GenerateBookDifferencesFile()
        //{
        //    Utils.SaveToXmlFile(PredefinedBookDifferences.RST, "G:\\rst.xml");
        //}


        private static void CreateSupplementalBible()
        {
            DateTime dtStart = DateTime.Now;

            string defaultNotebookFolderPath;
            OneNoteApp.GetSpecialLocation(SpecialLocation.slDefaultNotebookFolder, out defaultNotebookFolderPath);

            SupplementalBibleManager.CreateSupplementalBible(OneNoteApp, "kjv", defaultNotebookFolderPath, null);
            var result = SupplementalBibleManager.LinkSupplementalBibleWithMainBible(OneNoteApp, 0, null);

            DateTime dtEnd = DateTime.Now;

            var elapsed = dtEnd - dtStart;

            Console.WriteLine("Successfully! Elapsed time - {0} seconds", elapsed.TotalSeconds);

            int i = result.Errors.Count;
        }

        private static void GenerateParallelBible()
        {
            DateTime dtStart = DateTime.Now;

            var result = SupplementalBibleManager.AddParallelBible(OneNoteApp, "rst", null);

            DateTime dtEnd = DateTime.Now;

            var elapsed = dtEnd - dtStart;

            Console.WriteLine("Successfully! Elapsed time - {0} seconds", elapsed.TotalSeconds);

            int i = result.Errors.Count;
        }

        private static void ConvertRussianModule()
        {
            var converter = new BibleQuotaConverter("IBS", @"C:\Users\ademko\Dropbox\temp\IBS", @"c:\temp\IBS", Encoding.Default,
                "Ветхий Завет", "Новый Завет", 39, 27, "ru",
                PredefinedNotebooksInfo.Russian, PredefinedBookIndexes.RST, Utils.LoadFromXmlString<BibleTranslationDifferences>(Properties.Resources.rst), 
                "{0} глава. {1}", "2.0");

            converter.ConvertChapterNameFunc = (bookInfo, chapterNameInput) =>
            {
                int? chapterIndex = StringUtils.GetStringLastNumber(chapterNameInput);
                if (!chapterIndex.HasValue)
                    throw new Exception("Can not extract chapter index from string: " + chapterNameInput);
                return string.Format("{0} глава. {1}", chapterIndex, bookInfo.Name);
            };

            converter.Convert();

            int i = converter.Errors.Count;
        }

        private static void ConvertRomanModule()
        {
            var converter = new BibleQuotaConverter("Bible", @"C:\Temp\RCCV", @"c:\manifest.xml", Encoding.Unicode,
                "1. Vechiul Testament", "2. Noul Testament", 39, 27, "ro",
                PredefinedNotebooksInfo.English, PredefinedBookIndexes.KJV, new BibleTranslationDifferences(), "{0} capitolul. {1}", "2.0");

            converter.ConvertChapterNameFunc = (bookInfo, chapterNameInput) =>
            {
                int? chapterIndex = StringUtils.GetStringLastNumber(chapterNameInput);
                if (!chapterIndex.HasValue)
                    throw new Exception("Can not extract chapter index from string: " + chapterNameInput);
                return string.Format("{0} capitolul. {1}", chapterIndex, bookInfo.Name);
            };

            converter.Convert();
        }

        private static void ConvertEnglishModule()
        {
            var converter = new BibleQuotaConverter("KJV", @"C:\Temp\King_James_Version", @"c:\temp\KJV", Encoding.ASCII,
                "1. Old Testament", "2. New Testament", 39, 27, "en",
                PredefinedNotebooksInfo.English, PredefinedBookIndexes.KJV, new BibleTranslationDifferences(), "{0} chapter. {1}", "2.0");

            converter.ConvertChapterNameFunc = (bookInfo, chapterNameInput) =>
            {
                int? chapterIndex = StringUtils.GetStringLastNumber(chapterNameInput);
                if (!chapterIndex.HasValue)
                    throw new Exception("Can not extract chapter index from string: " + chapterNameInput);
                return string.Format("{0} chapter. {1}", chapterIndex, bookInfo.Name);
            };

            converter.Convert();
        }

        private static void SearchForEnText()
        {
            var oneNoteApp = new Microsoft.Office.Interop.OneNote.Application();
            string notebookId = OneNoteUtils.GetNotebookIdByName(oneNoteApp, "Biblia", false);

            var pages = OneNoteProxy.Instance.GetHierarchy(oneNoteApp, notebookId, Microsoft.Office.Interop.OneNote.HierarchyScope.hsPages, false);
            foreach (var page in pages.Content.Root.XPathSelectElements("//one:Page", pages.Xnm))
            {
                string pageId = page.Attribute("ID").Value;
                XmlNamespaceManager xnm;
                var pageDoc = OneNoteUtils.GetPageContent(oneNoteApp, pageId, out xnm);
                if (pageDoc.ToString().IndexOf("en-US") != -1)
                {
                    string pageName = page.Attribute("name").Value;
                }
            }
        }

        private static void ChangeCurrentPageLocale(string locale)
        {
            // it does not work (((((
            var oneNoteApp = new Microsoft.Office.Interop.OneNote.Application();

            XmlNamespaceManager xnm;
            var pageDoc = OneNoteUtils.GetPageContent(oneNoteApp, oneNoteApp.Windows.CurrentWindow.CurrentPageId, out xnm);

            foreach (var oe in pageDoc.Root.XPathSelectElements("//one:Cell", xnm))
            {
                if (oe.Attribute("lang") == null)
                    oe.Add(new XAttribute("lang", locale));
            }

            string pageXml = pageDoc.ToString();
            pageXml = pageXml.Replace("en-US", locale);

            oneNoteApp.UpdatePageContent(pageXml, DateTime.MinValue, Constants.CurrentOneNoteSchema);
        }

        private static void GenerateSummaryOfNotesNotebook()
        {
            NotebookGenerator.GenerateSummaryOfNotesNotebook(OneNoteApp, "Biblia", "Rezumatul de note");
        }

        private static void TryToUpdateInkNodes()
        {
            string xml;
            OneNoteApp.GetPageContent(OneNoteApp.Windows.CurrentWindow.CurrentPageId, out xml, Microsoft.Office.Interop.OneNote.PageInfo.piBasic, Constants.CurrentOneNoteSchema);

            System.Xml.XmlNamespaceManager xnm = new System.Xml.XmlNamespaceManager(new System.Xml.NameTable());
            var xd = System.Xml.Linq.XDocument.Parse(xml);

            xnm.AddNamespace("one", Constants.OneNoteXmlNs);
            System.Xml.Linq.XDocument doc = OneNoteUtils.GetXDocument(xml, out xnm);

            var inkNodes = doc.Root.XPathSelectElements("one:InkDrawing", xnm)
                             .Union(doc.Root.XPathSelectElements("one:Outline[.//one:InkWord]", xnm))
                //.Union(doc.Root.XPathSelectElements("//one:OE[.//one:InkDrawing]", xnm))
                             .ToArray();
            foreach (var inkNode in inkNodes)
                inkNode.Remove();


            //var oeInkNodes = doc.Root.XPathSelectElements("//one:OE[.//one:InkDrawing]", xnm).ToArray();
            //foreach (var oeInkNode in oeInkNodes)
            //{
            //    //var objectId = oeInkNode.Attribute("objectID").Value;
            //    var inkNode = oeInkNode.XPathSelectElement(".//one:InkDrawing", xnm);
            //   // inkNode.SetAttributeValue("objectID", objectId);
            //    doc.Root.Add(inkNode);
            //}


            //inkNodes = doc.Root.XPathSelectElements("//one:OE[.//one:InkDrawing]", xnm).ToArray();
            //foreach (var inkNode in inkNodes)
            //    inkNode.Remove();

            OneNoteApp.UpdatePageContent(doc.ToString(), DateTime.MinValue, Constants.CurrentOneNoteSchema);

        }

        
    }
}
