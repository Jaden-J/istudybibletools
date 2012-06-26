﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
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

namespace BibleConfigurator
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(params string[] args)
        {
            #region converter and test code
            try
            {
                //SearchForEnText();

                //ChangeCurrentPageLocale("ro");

                //TryToUpdateInkNodes();

                //ConvertEnglishModule();

                //ConvertRomanModule();

                //GenerateSummaryOfNotesNotebook();

                //DefaultRusModuleGenerator.GenerateModuleInfo();
            }
            catch (Exception ex)
            {
                Logger.LogError(ex.ToString());
            }

            #endregion

            try
            {
                LanguageManager.SetThreadUICulture();

                string message = BibleCommon.Resources.Constants.MoreThanSingleInstanceRun;
                if (args.Length == 1 && File.Exists(args[0]))
                    message += " " + BibleCommon.Resources.Constants.LoadMofuleInExistingInstance;

                FormExtensions.RunSingleInstance(message, () =>
                {
                    Application.EnableVisualStyles();
                    Application.SetCompatibleTextRenderingDefault(false);

                    Form form = PrepareForRunning(args);

                    if (form != null)
                    {                        
                        Application.Run(form);
                    }
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static Form PrepareForRunning(params string[] args)
        {
            Form result = null;

            try
            {                
                if (args.Contains(Consts.ShowModuleInfo) && SettingsManager.Instance.IsConfigured(OneNoteApp))
                    result = new AboutModuleForm(SettingsManager.Instance.ModuleName, true);
                else if (args.Contains(Consts.ShowAboutProgram))
                    result = new AboutProgramForm();
                else if (args.Contains(Consts.ShowManual))
                {
                    OpenManual();
                }
                else if (args.Contains(Consts.RunOnOneNoteStarts))
                {
                    if (SettingsManager.Instance.IsConfigured(OneNoteApp))
                    {
                        try
                        {
                            OneNoteLocker.LockAllBible(OneNoteApp);
                        }
                        catch (NotSupportedException)
                        {
                            //todo: log it
                        }
                    }
                    else
                        result = new MainForm(args);                    
                }
                else if (args.Contains(Consts.LockAllBible))
                {
                    try
                    {
                        OneNoteLocker.LockAllBible(OneNoteApp);
                    }
                    catch (NotSupportedException)
                    {
                        MessageBox.Show(BibleCommon.Resources.Constants.SkyDriveBibleIsNotSupportedForLock);
                    }
                }
                else if (args.Contains(Consts.UnlockAllBible))
                {
                    try
                    {
                        OneNoteLocker.UnlockAllBible(OneNoteApp);
                    }
                    catch (NotSupportedException)
                    {
                        MessageBox.Show(BibleCommon.Resources.Constants.SkyDriveBibleIsNotSupportedForLock);
                    }
                }
                else if (args.Contains(Consts.UnlockBibleSection))
                {
                    try
                    {
                        OneNoteLocker.UnlockCurrentSection(OneNoteApp);
                    }
                    catch (NotSupportedException)
                    {
                        MessageBox.Show(BibleCommon.Resources.Constants.SkyDriveBibleIsNotSupportedForLock);
                    }
                }
                else if (args.Length == 1)
                {
                    result = new MainForm(args);

                    if (!string.IsNullOrEmpty(args[0]))
                    {
                        string moduleFilePath = args[0];
                        if (File.Exists(moduleFilePath))
                        {

                            bool needToReload = ((MainForm)result).AddNewModule(moduleFilePath);
                            ((MainForm)result).ShowModulesTabAtStartUp = true;
                            ((MainForm)result).NeedToSaveChangesAfterLoadingModuleAtStartUp = needToReload;
                        }
                    }
                }
                else
                {   
                    result = new MainForm(args);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            if (_oneNoteApp != null)
                _oneNoteApp = null;

            return result;
        }

        public static void RunFromAnotherApp(params string[] args)
        {
            Form form = PrepareForRunning(args);

            if (form != null)
            {
                form.ShowDialog();
                form.Dispose();
            }
        }

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

  


        private static bool OpenManual()
        {
            var path = Path.GetDirectoryName(Path.GetDirectoryName(Utils.GetCurrentDirectory()));

            var files = Directory.GetFiles(path, string.Format("Instruction*{0}*", LanguageManager.UserLanguage.LCID));
            if (files.Length == 0)
                files = Directory.GetFiles(path, string.Format("Instruction*{0}*", LanguageManager.DefaultLCID));

            if (files.Length == 1)
            {
                Process.Start(files[0]);
                return true;
            }

            return false;
        }


       



        #region converter utils

        private static void ConvertRomanModule()
        {
            var converter = new BibleQuotaConverter("Bible", @"C:\Temp\RCCV", @"c:\manifest.xml", Encoding.Unicode,
                "1. Vechiul Testament", "2. Noul Testament", 39, 27, "ro", new List<NotebookInfo>()             
                {   
                    new NotebookInfo() { Type = NotebookType.Bible, Name = "Bible.onepkg" },
                    new NotebookInfo() { Type = NotebookType.BibleStudy, Name = "Bible Study.onepkg" },
                    new NotebookInfo() { Type = NotebookType.BibleComments, Name = "Comments to the Bible.onepkg" },
                    new NotebookInfo() { Type = NotebookType.BibleNotesPages, Name = "Summary of Notes.onepkg" }
                });

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
            var converter = new BibleQuotaConverter("Test", @"G:\Dropbox\Изучение Библии\программы\Цитата из Библии\King_James_Version", @"c:\manifest.xml", Encoding.ASCII,
                "1. Old Testament", "2. New Testament", 39, 27, null, new List<NotebookInfo>() 
                {  
                    new NotebookInfo() { Type = NotebookType.Bible, Name = "Bible.onepkg" },
                    new NotebookInfo() { Type = NotebookType.BibleStudy, Name = "Bible Study.onepkg" },
                    new NotebookInfo() { Type = NotebookType.BibleComments, Name = "Comments to the Bible.onepkg" },
                    new NotebookInfo() { Type = NotebookType.BibleNotesPages, Name = "Summary of Notes.onepkg" }
                });

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

            oneNoteApp.UpdatePageContent(pageXml);
        }

        private static void GenerateSummaryOfNotesNotebook()
        {
            NotebookGenerator.GenerateSummaryOfNotesNotebook("Biblia", "Rezumatul de note");
        }

        private static void TryToUpdateInkNodes()
        {
            var oneNoteApp = new Microsoft.Office.Interop.OneNote.Application();
            string xml;
            oneNoteApp.GetPageContent(oneNoteApp.Windows.CurrentWindow.CurrentPageId, out xml);

            System.Xml.XmlNamespaceManager xnm = new System.Xml.XmlNamespaceManager(new System.Xml.NameTable());
            var xd = System.Xml.Linq.XDocument.Parse(xml);

            xnm.AddNamespace("one", "http://schemas.microsoft.com/office/onenote/2010/onenote");
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

            oneNoteApp.UpdatePageContent(doc.ToString());
        }

        #endregion
    }
}