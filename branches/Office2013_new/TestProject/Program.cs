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
using BibleCommon.Scheme;
using TestProject.Properties;
using System.Text.RegularExpressions;


namespace TestProject
{    
    class Program
    {
        private const string ForGeneratingFolderPath = @"C:\Users\lux_demko\Desktop\temp\Dropbox\Holy Bible\IStudyBibleTools\ForGenerating";
        private const string TempFolderPath = @"C:\Users\lux_demko\Desktop\temp\temp";

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

       

        [STAThread]
        static void Main(string[] args)
        {
            Stopwatch sw = new Stopwatch();

            sw.Start();


            try
            {
                Assembly assembly = Assembly.LoadFrom(@"C:\Program Files\IStudyBibleTools\OneNote IStudyBibleTools\tools\BibleNoteLinker\BibleNoteLinker.exe");

                //GenerateBibleBooks();

                //SearchInNotebook();

                //TestModule();

                //CheckHTML();

                //GenerateBibleVersesLinks();

                //SearchStrongTerm(args);

                //AddColorLink();

                //GenerateDictionary();

                //GenerateStrongDictionary();
                
                //SearchForEnText();

                //ChangeCurrentPageLocale("ro");

                //TryToUpdateInkNodes();                

                //ConvertRussianModuleZefaniaXml();

                //ConvertEnglishModuleZefaniaXml();

                //ConvertUkrModule();

                //ConvertRussianModule();

                ConvertEnglishModule();

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

        private static void GenerateBibleBooks()
        {
            var manifestFilePath = @"C:\Users\lux_demko\Desktop\temp\Dropbox\manifest.xml";
            var bibleQuotaIniFilePath = "";
            var existingBooksFilePath = "";
            var targetFilePath = Path.Combine(TempFolderPath, "books.xml");

            
            var manifest = Utils.LoadFromXmlFile<ModuleInfo>(manifestFilePath);
            manifest.CorrectModuleAfterDeserialization();

            var booksInfo = new BibleBooksInfo()
            {
                Descr = manifest.ShortName,
                Alphabet = manifest.BibleStructure.Alphabet,
                ChapterString = "глава"
            };

            int bookIndex = 1;
            foreach (var oldBookInfo in manifest.BibleStructure.BibleBooks)
            {
                var bookInfo = new BookInfo()
                {
                    Index = bookIndex++,
                    Name = oldBookInfo.Name,
                    ShortNamesXMLString = string.Join(";", oldBookInfo.Abbreviations.Select(
                                abbr => !abbr.IsFullBookName 
                                            ? abbr.Value 
                                            : string.Format("'{0}'", abbr.Value)
                            ).ToArray())
                };

                booksInfo.Books.Add(bookInfo);
            }

            Utils.SaveToXmlFile(booksInfo, targetFilePath);
        }

        //private static void ConvertEnglishModuleZefaniaXml()
        //{
        //    string moduleName = "kjv";

        //    var converter = new ZefaniaXmlConverter(moduleName,
        //                                            "King James Version",
        //        Path.Combine(Path.Combine(ForGeneratingFolderPath, moduleName), BibleCommon.Consts.Constants.BibleInfoFileName),
        //        Utils.LoadFromXmlString<BibleBooksInfo>(Properties.Resources.BibleBooskInfo_kjv), Path.Combine(TempFolderPath, moduleName + "_zefaniaXml"), "en",
        //                                            PredefinedNotebooksInfo.English, new BibleTranslationDifferences(),  // вот эти тоже часто надо менять                
        //        "{0} chapter. {1}",
        //                                            PredefinedSectionsInfo.None, false, null, null,
        //        //PredefinedSectionsInfo.RSTStrong, true, "Стронга", 14700,   // параметры для стронга
        //        new Version(2, 0), false,
        //                                            ZefaniaXmlConverter.ReadParameters.None);  // и про эту не забыть

        //    converter.Convert();
        //}

        //private static void ConvertRussianModuleZefaniaXml()
        //{
        //    string moduleName = "rststrong";

        //    var converter = new ZefaniaXmlConverter(moduleName,
        //                                            "Русский синодальный перевод", 
        //        Path.Combine(Path.Combine(ForGeneratingFolderPath, moduleName), BibleCommon.Consts.Constants.BibleInfoFileName),               
        //        Utils.LoadFromXmlString<BibleBooksInfo>(Properties.Resources.BibleBooskInfo_rst), Path.Combine(TempFolderPath, moduleName + "_zefaniaXml"), "ru",
        //                                            PredefinedNotebooksInfo.Russian, Utils.LoadFromXmlString<BibleTranslationDifferences>(Properties.Resources.rst),  // вот эти тоже часто надо менять                
        //        "{0} глава. {1}",
        //                                            PredefinedSectionsInfo.None, false, null, null,
        //                                            //PredefinedSectionsInfo.RSTStrong, true, "Стронга", 14700,   // параметры для стронга
        //        new Version(2, 0), false,
        //                                            ZefaniaXmlConverter.ReadParameters.RemoveStrongs);  // и про эту не забыть

        //    converter.Convert();
        //}             

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
            //BibleVersesLinksCacheManager.GenerateBibleVersesLinks(OneNoteApp, SettingsManager.Instance.NotebookId_Bible, SettingsManager.Instance.SectionGroupId_Bible, new ConsoleLogger());            


          //  var result = BibleVersesLinksCacheManager.LoadBibleVersesLinks(SettingsManager.Instance.NotebookId_Bible);

            var verse = OneNoteProxy.Instance.GetVersePointerLink(new VersePointer("Фил 6"));
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
            var converter = new BibleQuotaDictionaryConverter(OneNoteApp, "Словари", "rststrong", "Словарь Стронга", "Еврейский и Греческий лексикон Стронга (с) Bob Jones University",
                new List<DictionaryFile>() { 
                    new DictionaryFile() { FilePath = Path.Combine(ForGeneratingFolderPath, @"Strongs\HEBREW.HTM"), SectionName = "Ветхий Завет.one", DictionaryPageDescription="Еврейский лексикон Стронга (с) Bob Jones University", TermPrefix = "H" },
                    new DictionaryFile() { FilePath = Path.Combine(ForGeneratingFolderPath, @"Strongs\GREEK.HTM"), SectionName = "Новый Завет.one", DictionaryPageDescription="Греческий лексикон Стронга (с) Bob Jones University", TermPrefix= "G" }
                }, BibleQuotaDictionaryConverter.StructureType.Strong, "Стронга",
                Path.Combine(TempFolderPath, "strong"), "<h4>", "Пользовательские заметки", "Найти все стихи с этим номером", "ru", new Version(2, 0));

            converter.Convert();

            using (var form = new ErrorsForm(converter.Errors.ConvertAll(er => er.Message)))
            {
                form.ShowDialog();
            }
        }

        private static void GenerateDictionary()
        {
            //var converter = new BibleQuotaDictionaryConverter(OneNoteApp, "Словари", "goetze", "Библейский словарь Б.Геце", "Библейский словарь Б.Геце",
            //  new List<DictionaryFile>() { 
            //        new DictionaryFile() { FilePath = Path.Combine(ForGeneratingFolderPath, @"Goetze\goetze.htm"), DictionaryPageDescription="Библейский словарь Б.Геце" }                    
            //    }, BibleQuotaDictionaryConverter.StructureType.Dictionary, "Геце",
            //    Path.Combine(TempFolderPath, "goetze"), "<h4>", "Пользовательские заметки", null, "ru", new Version(2, 0));


            var converter = new BibleQuotaDictionaryConverter(OneNoteApp, "Словари", "brockhaus", "Библейский словарь Брокгауза", "Библейский словарь Брокгауза",
             new List<DictionaryFile>() { 
                    new DictionaryFile() { FilePath = Path.Combine(ForGeneratingFolderPath, @"brockhaus\BrockhausLexicon.htm"), DictionaryPageDescription="Библейский словарь Брокгауза" }                    
                }, BibleQuotaDictionaryConverter.StructureType.Dictionary, "Брокгауза",
               Path.Combine(TempFolderPath, "brockhaus"), "<h4>", "Пользовательские заметки", null, "ru", new Version(2, 0));

            converter.Convert();

            using (var form = new ErrorsForm(converter.Errors.ConvertAll(er => er.Message)))
            {
                form.ShowDialog();
            }
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

            SupplementalBibleManager.CreateSupplementalBible(OneNoteApp, ModulesManager.GetModuleInfo("kjv"), defaultNotebookFolderPath, null);
            var result = SupplementalBibleManager.LinkSupplementalBibleWithPrimaryBible(OneNoteApp, 0, null, null);

            DateTime dtEnd = DateTime.Now;

            var elapsed = dtEnd - dtStart;

            Console.WriteLine("Successfully! Elapsed time - {0} seconds", elapsed.TotalSeconds);

            using (var form = new ErrorsForm(result.Errors.ConvertAll(er => er.Message)))
            {
                form.ShowDialog();
            }
        }

        private static void GenerateParallelBible()
        {
            DateTime dtStart = DateTime.Now;

            string defaultNotebookFolderPath;
            OneNoteApp.GetSpecialLocation(SpecialLocation.slDefaultNotebookFolder, out defaultNotebookFolderPath);

            var result = SupplementalBibleManager.AddParallelBible(OneNoteApp, ModulesManager.GetModuleInfo("rst"), null, null);

            DateTime dtEnd = DateTime.Now;

            var elapsed = dtEnd - dtStart;

            Console.WriteLine("Successfully! Elapsed time - {0} seconds", elapsed.TotalSeconds);

            using (var form = new ErrorsForm(result.Errors.ConvertAll(er => er.Message)))
            {
                form.ShowDialog();
            }
        }

        private static void ConvertUkrModule()
        {
            string moduleShortName = "UkrGYZ";
            var notebooksStructure = new NotebooksStructure() { Notebooks = PredefinedNotebooksInfo.Russian };  // это тоже часто меняется
            //notebooksStructure.DictionarySectionGroupName = "Стронга";  // параметры для стронга
            //notebooksStructure.DictionaryTermsCount = 14700;

            var converter = new BibleQuotaConverter(moduleShortName, Path.Combine(Path.Combine(ForGeneratingFolderPath, "old"), moduleShortName), Path.Combine(TempFolderPath, moduleShortName), "ru",
                notebooksStructure, PredefinedBookIndexes.RST, Utils.LoadFromXmlString<BibleTranslationDifferences>(Properties.Resources.rst),  // вот эти тоже часто надо менять                
                "{0} глава. {1}",
                false,                
                new Version(2, 0), false,
                BibleQuotaConverter.ReadParameters.None);  // и про эту не забыть

            converter.Convert();

            using (var form = new ErrorsForm(converter.Errors.ConvertAll(er => er.Message)))
            {
                form.ShowDialog();
            }
        }

        private static void ConvertRussianModule()
        {
            string moduleShortName = "rst77";
            var notebooksStructure = new NotebooksStructure() { Notebooks = PredefinedNotebooksInfo.Russian77 };  // это тоже часто меняется
            //notebooksStructure.DictionarySectionGroupName = "Стронга";  // параметры для стронга
            //notebooksStructure.DictionaryTermsCount = 14700;

            var converter = new BibleQuotaConverter(moduleShortName, Path.Combine(Path.Combine(ForGeneratingFolderPath, "old"), moduleShortName), Path.Combine(TempFolderPath, moduleShortName), "ru",
                notebooksStructure, PredefinedBookIndexes.RST77, Utils.LoadFromXmlString<BibleTranslationDifferences>(Properties.Resources.rst77),  // вот эти тоже часто надо менять                
                "{0} глава. {1}",
                false,
                //PredefinedSectionsInfo.RSTStrong, true, "Стронга", 14700,   // параметры для стронга
                new Version(2, 0), false,
                BibleQuotaConverter.ReadParameters.None);  // и про эту не забыть

            converter.Convert();

            using (var form = new ErrorsForm(converter.Errors.ConvertAll(er => er.Message)))
            {
                form.ShowDialog();
            }
        }

        private static void ConvertEnglishModule()
        {
            string moduleShortName = "dourh";
            var converter = new BibleQuotaConverter(moduleShortName, Path.Combine(Path.Combine(ForGeneratingFolderPath, "old"), moduleShortName), Path.Combine(TempFolderPath, moduleShortName),
                "en", new NotebooksStructure() { Notebooks = PredefinedNotebooksInfo.English }, PredefinedBookIndexes.KJV, new BibleTranslationDifferences(),
                "{0} chapter. {1}",
                false, 
                new Version(2, 0), false);

            converter.Convert();

            using (var form = new ErrorsForm(converter.Errors.ConvertAll(er => er.Message)))
            {
                form.ShowDialog();
            }
        }

        private static void ConvertRomanModule()
        {
            string moduleShortName = "rccv";
            var converter = new BibleQuotaConverter(moduleShortName, Path.Combine(Path.Combine(ForGeneratingFolderPath, "old"), moduleShortName), Path.Combine(TempFolderPath, moduleShortName),
                "ro", new NotebooksStructure() { Notebooks = PredefinedNotebooksInfo.English }, PredefinedBookIndexes.KJV, new BibleTranslationDifferences(),
                "{0} capitolul. {1}",
                false,
                new Version(2, 0), false);            

            converter.Convert();

            using (var form = new ErrorsForm(converter.Errors.ConvertAll(er => er.Message)))
            {
                form.ShowDialog();
            }
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
