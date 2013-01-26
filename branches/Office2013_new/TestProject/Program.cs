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
using System.Windows.Forms;


namespace TestProject
{    
    class Program
    {
        private const string ForGeneratingFolderPath = @"E:\Dropbox\Holy Bible\IStudyBibleTools\ForGenerating";
        private const string TempFolderPath = @"E:\temp";

        private static Microsoft.Office.Interop.OneNote.Application _oneNoteApp;   

        [STAThread]
        static void Main(string[] args)
        {
            Stopwatch sw = new Stopwatch();

            sw.Start();


            _oneNoteApp = new Microsoft.Office.Interop.OneNote.Application();

            try
            {

                //Console.WriteLine(Regex.Replace("<br>no<", string.Format("(^|[^0-9a-zA-Z]){0}($|[^0-9a-zA-Z<])", "no"), @"$1aeasdasds$2", RegexOptions.IgnoreCase));
                //return;

                //Console.WriteLine(StringUtils.GetQueryParameterValue("http://adjhdjkhsadsd.rudasd&sdsd=adsadasd&dsfsdf=sgfdsdfdsf&key=value", "key"));
                //return;

                //ConvertChineseModuleFromTextFiles();
                
                //GenerateBibleBooks();

                //SearchInNotebook();

                //TestModule();

                //CheckHTML();

                //GenerateBibleVersesLinks();

                //SearchStrongTerm(args);

                AddColorLink();

                //GenerateRuDictionary();

                //GenerateEnDictionary();

                //CorrectVineOT();

                //GenerateRuStrongDictionary();

                //GenerateEnStrongDictionary();
                
                //SearchForEnText();

                //ChangeCurrentPageLocale("ro");

                //TryToUpdateInkNodes();                                

                //ConvertUkrModule();

                //ConvertChineseModule();

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

        private static void CorrectVineOT()
        {
            var vineOTFilePath = Path.Combine(ForGeneratingFolderPath, string.Format("{0}\\{0}.htm", "vineot"));
            var strongOTFilePath = Path.Combine(ForGeneratingFolderPath, "rststrong\\HEBREW.HTM");

            var vineOT = File.ReadAllText(vineOTFilePath);
            var strongOT = File.ReadAllText(strongOTFilePath);

            var vineNumberSearchString = "<br>Strong's Number: ";
            var numberInHebrewSearchString = "<br>Original Word: <font face=\"BQTHeb\">";
            var endOfFontSearchString = "</font>";
            var numberInStrongInHebrewSearchString = "<span class=Index> <font face=\"BQTHeb\">";            
            var cursor = vineOT.IndexOf(vineNumberSearchString);
            var excludedNumbers = new int[] { 136 };

            var hebrewWordTemplate = "<font face=\"BQTHeb\">{0}</font>";
            var hebrewWordTemplate2 = "<font face=\"BQTHeb\"> {0}</font>";            

            while (cursor > -1)
            {
                var endOfLineIndex = vineOT.IndexOf(Environment.NewLine, cursor + 1);
                if (endOfLineIndex == -1)
                    throw new Exception("End of line is not found at " + cursor);


                var numberString = vineOT.Substring(cursor + vineNumberSearchString.Length, endOfLineIndex - cursor - vineNumberSearchString.Length);

                if (numberString.IndexOf(",") != -1)
                {
                    Console.WriteLine(numberString);
                    numberString = numberString.Split(new char[] { ',' })[0];
                }

                int number;
                if (!int.TryParse(numberString, out number))
                    throw new Exception("Can not parse " + numberString);

                if (!excludedNumbers.Contains(number))
                {

                    var numberInHebrewIndex = vineOT.IndexOf(numberInHebrewSearchString, endOfLineIndex);
                    if (numberInHebrewIndex == -1 || numberInHebrewIndex - endOfLineIndex > 5)
                        throw new Exception("Can not find numberInHebrewIndex for " + number);

                    var numberInHebrewEndIndex = vineOT.IndexOf(endOfFontSearchString, numberInHebrewIndex);
                    if (numberInHebrewEndIndex == -1)
                        throw new Exception("Can not find numberInHebrewEndIndex for " + number);

                    var hebrewWordInVine = vineOT.Substring(numberInHebrewIndex + numberInHebrewSearchString.Length, numberInHebrewEndIndex - numberInHebrewIndex - numberInHebrewSearchString.Length).Trim();

                    var numberInStrongSearchString = string.Format("<h4>{0:00000}</h4>", number);
                    var numberInStrongIndex = strongOT.IndexOf(numberInStrongSearchString);
                    if (numberInStrongIndex == -1)
                        throw new Exception("Strong number can not be found in strongOT: " + number);

                    var numberInStrongInHebrewIndex = strongOT.IndexOf(numberInStrongInHebrewSearchString, numberInStrongIndex);
                    if (numberInStrongInHebrewIndex == -1)
                        throw new Exception("Can not find numberInStrongInHebrewIndex for " + number);

                    var numberInStrongInHebrewEndIndex = strongOT.IndexOf(endOfFontSearchString, numberInStrongInHebrewIndex);
                    if (numberInStrongInHebrewEndIndex == -1)
                        throw new Exception("Can not find numberInStrongInHebrewEndIndex for " + number);

                    var hebrewWordInStrong = strongOT.Substring(numberInStrongInHebrewIndex + numberInStrongInHebrewSearchString.Length, numberInStrongInHebrewEndIndex - numberInStrongInHebrewIndex - numberInStrongInHebrewSearchString.Length).Trim();
                    vineOT = vineOT.Replace(hebrewWordInVine, hebrewWordInStrong);

                    vineOT = vineOT.Replace(string.Format(hebrewWordTemplate, hebrewWordInVine), string.Format(hebrewWordTemplate, hebrewWordInStrong));
                    vineOT = vineOT.Replace(string.Format(hebrewWordTemplate2, hebrewWordInVine), string.Format(hebrewWordTemplate, hebrewWordInStrong));                    

                    //vineOT = ReplaceEx(vineOT, hebrewWordInVine, string.Format(hebrewWordTemplate, hebrewWordInStrong));

                    Regex rgx = new Regex(string.Format("(^|[^0-9a-zA-Z]){0}($|[^0-9a-zA-Z<])",
                                hebrewWordInVine
                                    .Replace(@"\", @"\\")
                                    .Replace("(", @"\(")
                                    .Replace(")", @"\)")
                                    .Replace("[", @"\[")
                                    .Replace("]", @"\]")),
                                    RegexOptions.IgnoreCase);
                    vineOT = rgx.Replace(vineOT, string.Format("$1" + hebrewWordTemplate + "$2", hebrewWordInStrong));
                }

                cursor = vineOT.IndexOf(vineNumberSearchString, cursor + vineNumberSearchString.Length + 1);
            }

            File.WriteAllText(vineOTFilePath + "_new", vineOT);
        }



        private static void ConvertChineseModuleFromTextFiles()
        {
            var folder = @"C:\TEMP\temp\ncv-t";
            var converter = new TextFilesConverter(folder, Path.Combine(TempFolderPath, Path.GetFileName(folder)));
            converter.Convert();

            using (var form = new ErrorsForm(converter.Errors.ConvertAll(er => er.Message)))
            {
                form.ShowDialog();
            }
        }              

        private static void GenerateBibleBooks()
        {
            var manifestFilePath = @"C:\Users\lux_demko\Desktop\temp\Dropbox\manifest.xml";            
            var targetFilePath = Path.Combine(TempFolderPath, "books.xml");

            
            var manifest = Utils.LoadFromXmlFile<ModuleInfo>(manifestFilePath);
            manifest.CorrectModuleAfterDeserialization();

            var booksInfo = new BibleBooksInfo()
            {
                Descr = manifest.ShortName,
                Alphabet = manifest.BibleStructure.Alphabet,
                ChapterPageNameTemplate = "{0} глава. {1}"
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
            string xml;
            var xnm = new XmlNamespaceManager(new NameTable());
            xnm.AddNamespace("one", "http://schemas.microsoft.com/office/onenote/2013/onenote");
            var oneNoteApp = new Microsoft.Office.Interop.OneNote.Application();

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
            var pageDoc = OneNoteUtils.GetPageContent(ref _oneNoteApp, _oneNoteApp.Windows.CurrentWindow.CurrentPageId, out xnm);

            string s = @"";

            NotebookGenerator.AddTextElementToPage(pageDoc, s);
            
            OneNoteUtils.UpdatePageContentSafe(ref _oneNoteApp, pageDoc, xnm);
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
            oneNoteApp.GetPageContent(_oneNoteApp.Windows.CurrentWindow.CurrentPageId, out xml);
            var currentPageDoc = XDocument.Parse(xml);

            var nms = XNamespace.Get("http://schemas.microsoft.com/office/onenote/2013/onenote");

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

        private static void GenerateEnStrongDictionary()
        {
            var moduleName = "kjvstrong";
            var converter = new BibleQuotaDictionaryConverter(_oneNoteApp, "Dictionaries", moduleName, "Strong's Dictionary", "Strong's Exhaustive Concordance (c) Bible Foundation",
                 new List<DictionaryFile>() { 
                    new DictionaryFile() { FilePath = Path.Combine(ForGeneratingFolderPath, moduleName + "\\HEBREW.HTM"), SectionName = "1. Old Testament.one", DictionaryPageDescription="Strong's Hebrew Dictionary (с) Bible Foundation", TermPrefix = "H" },
                    new DictionaryFile() { FilePath = Path.Combine(ForGeneratingFolderPath, moduleName + "\\GREEK.HTM"), SectionName = "2. New Testament.one", DictionaryPageDescription="Strong's Greek Dictionary (с) Bible Foundation", TermPrefix= "G" }
                }, BibleQuotaDictionaryConverter.StructureType.Strong, "Strong",
                 Path.Combine(TempFolderPath, moduleName), "<h4>", "User Notes", "Find all verses with this number", "en", new Version(2, 0));

            converter.Convert();

            using (var form = new ErrorsForm(converter.Errors.ConvertAll(er => er.Message)))
            {
                form.ShowDialog();
            }

            MessageBox.Show(string.Format("Не забыть обновить в {0}.structure.xml аттрибуты: XmlDictionaryPagesCount и XmlDictionaryTermsCount", moduleName));                
        }

        private static void GenerateRuStrongDictionary()
        {
            var moduleName = "rststrong";

            var converter = new BibleQuotaDictionaryConverter(_oneNoteApp, "Словари", moduleName, "Словарь Стронга", "Еврейский и Греческий лексикон Стронга (с) Bob Jones University",
                new List<DictionaryFile>() { 
                    new DictionaryFile() { FilePath = Path.Combine(ForGeneratingFolderPath, moduleName + "\\HEBREW.HTM"), SectionName = "Ветхий Завет.one", DictionaryPageDescription="Еврейский лексикон Стронга (с) Bob Jones University", TermPrefix = "H" },
                    new DictionaryFile() { FilePath = Path.Combine(ForGeneratingFolderPath, moduleName + "\\GREEK.HTM"), SectionName = "Новый Завет.one", DictionaryPageDescription="Греческий лексикон Стронга (с) Bob Jones University", TermPrefix= "G" }
                }, BibleQuotaDictionaryConverter.StructureType.Strong, "Стронга",
                Path.Combine(TempFolderPath, moduleName), "<h4>", "Пользовательские заметки", "Найти все стихи с этим номером", "ru", new Version(2, 0));

            converter.Convert();

            using (var form = new ErrorsForm(converter.Errors.ConvertAll(er => er.Message)))
            {
                form.ShowDialog();
            }

            MessageBox.Show(string.Format("Не забыть обновить в {0}.structure.xml аттрибуты: XmlDictionaryPagesCount и XmlDictionaryTermsCount", moduleName));                
        }

        private static void GenerateRuDictionary()
        {
            //var converter = new BibleQuotaDictionaryConverter(OneNoteApp, "Словари", "goetze", "Библейский словарь Б.Геце", "Библейский словарь Б.Геце",
            //  new List<DictionaryFile>() { 
            //        new DictionaryFile() { FilePath = Path.Combine(ForGeneratingFolderPath, @"Goetze\goetze.htm"), DictionaryPageDescription="Библейский словарь Б.Геце" }                    
            //    }, BibleQuotaDictionaryConverter.StructureType.Dictionary, "Геце",
            //    Path.Combine(TempFolderPath, "goetze"), "<h4>", "Пользовательские заметки", null, "ru", new Version(2, 0));


            var converter = new BibleQuotaDictionaryConverter(_oneNoteApp, "Словари", "brockhaus", "Библейский словарь Брокгауза", "Библейский словарь Брокгауза",
             new List<DictionaryFile>() { 
                    new DictionaryFile() { FilePath = Path.Combine(ForGeneratingFolderPath, @"brockhaus\BrockhausLexicon.htm"), DictionaryPageDescription="Библейский словарь Брокгауза" }                    
                }, BibleQuotaDictionaryConverter.StructureType.Dictionary, "Брокгауза",
               Path.Combine(TempFolderPath, "brockhaus"), "<h4>", "Пользовательские заметки", null, "ru", new Version(2, 0));


            using (var form = new ErrorsForm(converter.Errors.ConvertAll(er => er.Message)))
            {
                form.ShowDialog();
            }
        }

        private static void GenerateEnDictionary()
        {
            var moduleName = "vinent";
            var moduleDescription = "Vine's Expository Dictionary of New Testament Words";
            var converter = new BibleQuotaDictionaryConverter(_oneNoteApp, "Dictionaries", moduleName, moduleDescription, moduleDescription,
             new List<DictionaryFile>() { 
                    new DictionaryFile() { FilePath = Path.Combine(ForGeneratingFolderPath, string.Format("{0}\\{0}.htm", moduleName)), DictionaryPageDescription = moduleDescription }
                }, BibleQuotaDictionaryConverter.StructureType.Dictionary, moduleName,
              Path.Combine(TempFolderPath, moduleName), "<p><b>", "User Notes", null, "en", new Version(2, 0));

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

            string defaultNotebookFolderPath = null;
            OneNoteUtils.UseOneNoteAPI(ref _oneNoteApp, () =>
            {
                _oneNoteApp.GetSpecialLocation(SpecialLocation.slDefaultNotebookFolder, out defaultNotebookFolderPath);
            });

            SupplementalBibleManager.CreateSupplementalBible(ref _oneNoteApp, ModulesManager.GetModuleInfo("kjv"), defaultNotebookFolderPath, null);
            var result = SupplementalBibleManager.LinkSupplementalBibleWithPrimaryBible(ref _oneNoteApp, null, null);

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

            string defaultNotebookFolderPath = null;
            OneNoteUtils.UseOneNoteAPI(ref _oneNoteApp, () =>
            {
                _oneNoteApp.GetSpecialLocation(SpecialLocation.slDefaultNotebookFolder, out defaultNotebookFolderPath);
            });

            var result = SupplementalBibleManager.AddParallelBible(ref _oneNoteApp, ModulesManager.GetModuleInfo("rst"), null, null);

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
            string moduleShortName = "ukr";
            var notebooksStructure = new NotebooksStructure() { Notebooks = PredefinedNotebooksInfo.Russian };  // это тоже часто меняется
            //notebooksStructure.DictionarySectionGroupName = "Стронга";  // параметры для стронга
            //notebooksStructure.DictionaryTermsCount = 14700;

            var converter = new BibleQuotaConverter(moduleShortName, Path.Combine(Path.Combine(ForGeneratingFolderPath, "old"), moduleShortName), Path.Combine(TempFolderPath, moduleShortName), "uk",
                notebooksStructure, PredefinedBookIndexes.KJV, Utils.LoadFromXmlString<BibleTranslationDifferences>(Properties.Resources.rst),  // вот эти тоже часто надо менять                
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

        private static void ConvertChineseModule()
        {
            string moduleShortName = "cuvs";
            var converter = new BibleQuotaConverter(moduleShortName, Path.Combine(Path.Combine(ForGeneratingFolderPath, "old"), moduleShortName), Path.Combine(TempFolderPath, moduleShortName),
                "zh_CN", new NotebooksStructure() { Notebooks = PredefinedNotebooksInfo.English }, PredefinedBookIndexes.KJV, new BibleTranslationDifferences(),
                "{0} chapter. {1}",
                false,
                new Version(2, 0), false);

            converter.Convert();

            using (var form = new ErrorsForm(converter.Errors.ConvertAll(er => er.Message)))
            {
                form.ShowDialog();
            }
        }      

        private static void ConvertEnglishModule()
        {
            string moduleShortName = "kjvstrong";

            var notebooksStructure = new NotebooksStructure() { Notebooks = PredefinedNotebooksInfo.English };
            notebooksStructure.DictionarySectionGroupName = "Strong";
            notebooksStructure.DictionaryTermsCount = 14198;
            notebooksStructure.DictionaryPagesCount = 141;
            var converter = new BibleQuotaConverter(moduleShortName, Path.Combine(Path.Combine(ForGeneratingFolderPath, "old"), moduleShortName), Path.Combine(TempFolderPath, moduleShortName),
                "en", notebooksStructure, PredefinedBookIndexes.KJV, new BibleTranslationDifferences(),
                "{0} chapter. {1}",
                true, 
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
            string notebookId = OneNoteUtils.GetNotebookIdByName(ref oneNoteApp, "Biblia", false);

            var pages = OneNoteProxy.Instance.GetHierarchy(ref oneNoteApp, notebookId, Microsoft.Office.Interop.OneNote.HierarchyScope.hsPages, false);
            foreach (var page in pages.Content.Root.XPathSelectElements("//one:Page", pages.Xnm))
            {
                string pageId = (string)page.Attribute("ID");
                XmlNamespaceManager xnm;
                var pageDoc = OneNoteUtils.GetPageContent(ref oneNoteApp, pageId, out xnm);
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
            var pageDoc = OneNoteUtils.GetPageContent(ref oneNoteApp, oneNoteApp.Windows.CurrentWindow.CurrentPageId, out xnm);

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
            _oneNoteApp.GetPageContent(_oneNoteApp.Windows.CurrentWindow.CurrentPageId, out xml, Microsoft.Office.Interop.OneNote.PageInfo.piBasic, Constants.CurrentOneNoteSchema);

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

            _oneNoteApp.UpdatePageContent(doc.ToString(), DateTime.MinValue, Constants.CurrentOneNoteSchema);

        }

        
    }
}
