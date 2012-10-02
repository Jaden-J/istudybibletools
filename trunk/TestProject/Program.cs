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
using BibleCommon.UI.Forms;
using BibleCommon.Handlers;
using System.Xml.Serialization;


namespace TestProject
{
    class Program
    {
        private const string ForGeneratingFolderPath = @"C:\Users\lux_demko\Desktop\temp\Dropbox\Holy Bible\ForGenerating\";
        private const string TempFolderPath = @"C:\Users\lux_demko\Desktop\temp\temp\";

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
            Stopwatch sw = new Stopwatch();

            sw.Start();

            try
            {               
                //SearchInNotebook();

                //TestModule();

                //CheckHTML();

                //GenerateBibleVersesLinks();

                //SearchStrongTerm(args);

                //AddColorLink();

                GenerateDictionary();

                //GenerateStrongDictionary();
                
                //SearchForEnText();

                //ChangeCurrentPageLocale("ro");

                //TryToUpdateInkNodes();

                //ConvertRussianModule();

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

            sw.Stop();

            Console.WriteLine("Finish. Elapsed time: {0}", sw.Elapsed);
            Console.ReadKey();
        }

        private static void SearchInNotebook()
        {
            var xnm = new XmlNamespaceManager(new NameTable());
            xnm.AddNamespace("one", Constants.OneNoteXmlNs);

            var oneNoteApp = new Microsoft.Office.Interop.OneNote.Application();
            string xml;
            oneNoteApp.GetHierarchy(null, HierarchyScope.hsNotebooks, out xml);
            var notebookId = (string)XDocument.Parse(xml).Root.XPathSelectElement("one:Notebook", xnm).Attribute("ID");
            oneNoteApp.FindPages(notebookId, "test", out xml, true, true);
        }      

        private static void TestModule()
        {
            string filePath = @"C:\Users\lux_demko\Desktop\temp\Dropbox\temp\Modules\RST\manifest.xml";
            var _serializer = new XmlSerializer(typeof(ModuleInfo));

            using (var fs = new FileStream(filePath, FileMode.Open))
            {
                var module = (ModuleInfo)_serializer.Deserialize(fs);
                module.CorrectModuleAfterDeserialization();
                var d = module.BibleStructure;
            }             
        }

        private static void CheckHTML()
        {
            XmlNamespaceManager xnm;
            var pageDoc = OneNoteUtils.GetPageContent(OneNoteApp, OneNoteApp.Windows.CurrentWindow.CurrentPageId, out xnm);

            string s = @"";

            NotebookGenerator.AddTextElementToPage(pageDoc, s);

            OneNoteUtils.UpdatePageContentSafe(OneNoteApp, pageDoc, xnm);
        }

        private static void GenerateBibleVersesLinks()
        {
            Stopwatch sw = new Stopwatch();
            //sw.Start();
            //BibleVersesLinksCacheManager.GenerateBibleVersesLinks(OneNoteApp, SettingsManager.Instance.NotebookId_Bible, SettingsManager.Instance.SectionGroupId_Bible, null);
            //sw.Stop();
            //Console.WriteLine(sw.Elapsed.TotalSeconds);

            sw.Start();
            var result = BibleVersesLinksCacheManager.LoadBibleVersesLinks(SettingsManager.Instance.NotebookId_Bible);
            sw.Stop();
            Console.WriteLine(sw.Elapsed.TotalSeconds);
        }

        private static void SearchStrongTerm(string[] args)
        {
            var handler = new FindVersesWithStrongNumberHandler();
            handler.ExecuteCommand(args);
        }

        private static void AddColorLink()
        {   
            var oneNoteApp = new Microsoft.Office.Interop.OneNote.Application();             
            string xml;
            oneNoteApp.GetPageContent(OneNoteApp.Windows.CurrentWindow.CurrentPageId, out xml);
            var currentPageDoc = XDocument.Parse(xml);

            var nms = XNamespace.Get("http://schemas.microsoft.com/office/onenote/2010/onenote");

            var textEl = new XElement(nms + "Outline",
                            new XElement(nms + "OEChildren",
                                new XElement(nms + "OE",
                                    new XElement(nms + "T",
                                        new XCData(                                            
                                            "text before <a href='http://google.com'><span style='color:#000000'>test link</span></a> text after"
                                            )))));

            currentPageDoc.Root.Add(textEl);
            oneNoteApp.UpdatePageContent(currentPageDoc.ToString());
        }

        private static void GenerateStrongDictionary()
        {
            var converter = new BibleQuotaDictionaryConverter(OneNoteApp, "Словари", "Strong", "Словарь Стронга",
                new List<DictionaryFile>() { 
                    new DictionaryFile() { FilePath = Path.Combine(ForGeneratingFolderPath, @"Strongs2\HEBREW.HTM"), SectionName = "Ветхий Завет.one", DisplayName="Еврейский лексикон Стронга (с) Bob Jones University", TermPrefix = "H", StartIndex = 0 },
                    new DictionaryFile() { FilePath = Path.Combine(ForGeneratingFolderPath, @"Strongs2\GREEK.HTM"), SectionName = "Новый Завет.one", DisplayName="Греческий лексикон Стронга (с) Bob Jones University", TermPrefix= "G", StartIndex = 0 }
                }, BibleQuotaDictionaryConverter.StructureType.Strong, "Стронга",
                Path.Combine(TempFolderPath, "strong"), "<h4>", "Пользовательские заметки", "Найти все стихи с этим номером", "ru", "2.0");

            converter.Convert();

            var form = new ErrorsForm(converter.Errors.ConvertAll(er => er.Message));
            form.ShowDialog();           
        }

        private static void GenerateDictionary()
        {
            var converter = new BibleQuotaDictionaryConverter(OneNoteApp, "Словари", "Brockhaus", "Библейский словарь Брокгауза",
              new List<DictionaryFile>() { 
                    new DictionaryFile() { FilePath = Path.Combine(ForGeneratingFolderPath, @"Brockhaus\BrockhausLexicon.htm"), SectionName = "Брокгауза.one", DisplayName="Библейский словарь Брокгауза" }                    
                }, BibleQuotaDictionaryConverter.StructureType.Dictionary, null,
                Path.Combine(TempFolderPath, "Brockhaus"), "<h4>", "Пользовательские заметки", null, "ru", "2.0");

            converter.Convert();

            var form = new ErrorsForm(converter.Errors.ConvertAll(er => er.Message));
            form.ShowDialog();        
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
            var result = SupplementalBibleManager.LinkSupplementalBibleWithMainBible(OneNoteApp, 0, null, null);

            DateTime dtEnd = DateTime.Now;

            var elapsed = dtEnd - dtStart;

            Console.WriteLine("Successfully! Elapsed time - {0} seconds", elapsed.TotalSeconds);

            var form = new ErrorsForm(result.Errors.ConvertAll(er => er.Message));
            form.ShowDialog();           
        }

        private static void GenerateParallelBible()
        {
            DateTime dtStart = DateTime.Now;

            string defaultNotebookFolderPath;
            OneNoteApp.GetSpecialLocation(SpecialLocation.slDefaultNotebookFolder, out defaultNotebookFolderPath);

            var result = SupplementalBibleManager.AddParallelBible(OneNoteApp, "rst", defaultNotebookFolderPath, new Dictionary<string,string>(), null);

            DateTime dtEnd = DateTime.Now;

            var elapsed = dtEnd - dtStart;

            Console.WriteLine("Successfully! Elapsed time - {0} seconds", elapsed.TotalSeconds);

            var form = new ErrorsForm(result.Errors.ConvertAll(er => er.Message));
            form.ShowDialog();           
        }

        private static void ConvertRussianModule()
        {
            string moduleShortName = "rst";
            var converter = new BibleQuotaConverter(moduleShortName, Path.Combine(ForGeneratingFolderPath, moduleShortName), Path.Combine(TempFolderPath, moduleShortName), 
                "ru", PredefinedNotebooksInfo.Russian, PredefinedBookIndexes.RST, Utils.LoadFromXmlString<BibleTranslationDifferences>(Properties.Resources.rst), "{0} глава. {1}",
                null, 
                false, null, null, // параметры для стронга
                "2.0", true);            

            converter.Convert();

            var form = new ErrorsForm(converter.Errors.ConvertAll(er => er.Message));
            form.ShowDialog();           
        }

        private static void ConvertEnglishModule()
        {
            string moduleShortName = "kjv";
            var converter = new BibleQuotaConverter(moduleShortName, Path.Combine(ForGeneratingFolderPath, moduleShortName), Path.Combine(TempFolderPath, moduleShortName),
                "en", PredefinedNotebooksInfo.English, PredefinedBookIndexes.KJV, new BibleTranslationDifferences(),
                "{0} chapter. {1}",
                null,
                false, null, null, // параметры для стронга
                "2.0", true);

            converter.Convert();

            var form = new ErrorsForm(converter.Errors.ConvertAll(er => er.Message));
            form.ShowDialog();
        }

        private static void ConvertRomanModule()
        {
            string moduleShortName = "rccv";
            var converter = new BibleQuotaConverter(moduleShortName, Path.Combine(ForGeneratingFolderPath, moduleShortName), Path.Combine(TempFolderPath, moduleShortName), 
                "ro", PredefinedNotebooksInfo.English, PredefinedBookIndexes.KJV, new BibleTranslationDifferences(),
                "{0} capitolul. {1}",
                null, 
                false, null, null, 
                "2.0", false);            

            converter.Convert();

            var form = new ErrorsForm(converter.Errors.ConvertAll(er => er.Message));
            form.ShowDialog();           
        }       

        private static void SearchForEnText()
        {
            var oneNoteApp = new Microsoft.Office.Interop.OneNote.Application();
            string notebookId = OneNoteUtils.GetNotebookIdByName(oneNoteApp, "Biblia", false);

            var pages = OneNoteProxy.Instance.GetHierarchy(oneNoteApp, notebookId, Microsoft.Office.Interop.OneNote.HierarchyScope.hsPages, false);
            foreach (var page in pages.Content.Root.XPathSelectElements("//one:Page", pages.Xnm))
            {
                string pageId = (string)page.Attribute("ID");
                XmlNamespaceManager xnm;
                var pageDoc = OneNoteUtils.GetPageContent(oneNoteApp, pageId, out xnm);
                if (pageDoc.ToString().IndexOf("en-US") != -1)
                {
                    string pageName = (string)page.Attribute("name");
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
