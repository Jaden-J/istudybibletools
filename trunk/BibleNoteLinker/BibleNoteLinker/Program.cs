using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.OneNote;
using System.Xml.Linq;
using System.Xml.XPath;
using BibleCommon;
using System.Xml;
using System.Configuration;
using BibleCommon.Services;
using BibleCommon.Helpers;
using System.IO;
using System.Diagnostics;

namespace BibleNoteLinker
{
    public class Program
    {
        const string Arg_AllPages = "-allpages";
        const string Arg_DeleteNotes = "-deletenotes";
        const string Arg_Force = "-force";        

        public class Args
        {
            public bool AnalyzeAllPages { get; set; }  // false - значит только текущую
            public NoteLinkManager.AnalyzeDepth AnalyzeDepth { get; set; }
            public bool Force { get; set; }            
            public bool DeleteNotes { get; set; }

            public Args()
            {
                this.AnalyzeAllPages = false;
                this.AnalyzeDepth = NoteLinkManager.AnalyzeDepth.Full;
                this.Force = false;
                this.DeleteNotes = false;
            }
        }

        private static Args GetUserArgs(string[] args)
        {            
            Args result = new Args();

            for (int i = 0; i < args.Length; i++)
            {   
                string argLower = args[i].ToLower();
                if (argLower == Arg_AllPages)
                    result.AnalyzeAllPages = true;
                else if (argLower == Arg_Force)
                    result.Force = true;
                else if (argLower == Arg_DeleteNotes)
                    result.DeleteNotes = true;                
                else
                {
                    int temp;
                    if (int.TryParse(argLower, out temp))
                        result.AnalyzeDepth = (NoteLinkManager.AnalyzeDepth)temp;
                }
            }

            return result;
        }

        static void Main(string[] args)
        {
            Logger.Init("BibleNoteLinker");
            DateTime dtStart = DateTime.Now;            

            try
            {
                Logger.LogMessage("Время старта: {0}", dtStart.ToLongTimeString());

                Args userArgs = GetUserArgs(args);
                if (userArgs.DeleteNotes)
                    Logger.LogMessage("Удаляем страницы заметок и ссылки на них.");
                else
                {
                    Logger.LogMessage("Уровень текущего анализа: '{0} ({1})'.", userArgs.AnalyzeDepth, (int)userArgs.AnalyzeDepth);
                    if (userArgs.Force)
                        Logger.LogMessage("Анализируем ссылки в том числе.");
                }

                Application oneNoteApp = new Application();                               

                if (userArgs.AnalyzeAllPages)
                {
                    Logger.LogMessage("Старт обработки всей записной книжки");

                    if (userArgs.DeleteNotes)
                        ProcessNotebook(oneNoteApp, Settings.Default.NotebookName_Bible, Settings.Default.SectionGroupName_Bible, userArgs);
                    else
                    {
                        ProcessNotebook(oneNoteApp, Settings.Default.NotebookName_BibleComments, Settings.Default.SectionGroupName_BibleComments, userArgs);
                        ProcessNotebook(oneNoteApp, Settings.Default.NotebookName_BibleStudy, Settings.Default.SectionGroupName_BibleStudy, userArgs);
                    }
                }
                else
                {
                    Logger.LogMessage("Старт обработки текущей страницы");

                    if (oneNoteApp.Windows.CurrentWindow != null)
                    {
                        string currentPageId = oneNoteApp.Windows.CurrentWindow.CurrentPageId;
                        string currentSectionId = oneNoteApp.Windows.CurrentWindow.CurrentSectionId;
                        string currentSectionGroupId = oneNoteApp.Windows.CurrentWindow.CurrentSectionGroupId;
                        string currentNotebookId = oneNoteApp.Windows.CurrentWindow.CurrentNotebookId;

                        if (!string.IsNullOrEmpty(currentPageId))
                        {
                            if (userArgs.DeleteNotes)
                                NoteLinkManager.DeletePageNotes(oneNoteApp, currentNotebookId, currentSectionGroupId, currentSectionId, currentPageId, OneNoteUtils.GetHierarchyElementName(oneNoteApp,currentPageId));
                            else
                                NoteLinkManager.LinkPageVerses(oneNoteApp, currentNotebookId, currentSectionGroupId, currentSectionId, currentPageId, userArgs.AnalyzeDepth, userArgs.Force);
                        }
                        else
                            Logger.LogError("Не найдено открытой страницы заметок");
                    }
                    else
                    {
                        Logger.LogError("Не найдено открытой записной книжки");
                    }
                }

                Logger.LogMessage("Успешно завершено.");                
            }
            catch (Exception ex)
            {
                Logger.LogError(ex);
            }

            Logger.LogMessage("Времени затрачено: {0}", DateTime.Now.Subtract(dtStart));

            Logger.Done();

            if (Logger.ErrorWasLogged)
            {
                Console.WriteLine("Во время работы программы произошли ошибки.");
                Console.ReadKey();
            }
        }

        private static void ProcessNotebook(Application oneNoteApp, string notebookName, string sectionGroupName, Args userArgs)
        {
            string notebookId = OneNoteUtils.GetNotebookId(oneNoteApp, notebookName);

            if (!string.IsNullOrEmpty(notebookId))
            {
                Logger.LogMessage("Обработка записной книжки: '{0}'", OneNoteUtils.GetHierarchyElementName(oneNoteApp, notebookId));  // чтобы точно убедиться

                string hierarchyXml;
                oneNoteApp.GetHierarchy(notebookId, HierarchyScope.hsPages, out hierarchyXml);
                XmlNamespaceManager xnm;
                XDocument notebookDoc = OneNoteUtils.GetXDocument(hierarchyXml, out xnm);

                Logger.MoveLevel(1);
                ProcessRootSectionGroup(oneNoteApp, notebookId, notebookDoc, sectionGroupName, xnm, userArgs.AnalyzeDepth, userArgs.Force, userArgs.DeleteNotes);
                Logger.MoveLevel(-1);
            }
            else
            {
                Logger.LogError(string.Format("Не найдено записной книжки '{0}'.", notebookName));
            }
        }

        private static void ProcessRootSectionGroup(Application oneNoteApp, string notebookId, XDocument doc, string sectionGroupName,
            XmlNamespaceManager xnm, NoteLinkManager.AnalyzeDepth linkDepth, bool force, bool deleteNotes)
        {
            XElement sectionGroup = string.IsNullOrEmpty(sectionGroupName) 
                                        ? doc.Root 
                                        : doc.Root.XPathSelectElement(
                                                string.Format("one:SectionGroup[@name='{0}']", sectionGroupName), xnm);                        

            if (sectionGroup != null)
                ProcessSectionGroup(sectionGroup, sectionGroupName, oneNoteApp, notebookId, xnm, linkDepth, force, deleteNotes);
            else
                Logger.LogError("Не удаётся найти группу секций '{0}'", sectionGroupName);
        }

        private static void ProcessSectionGroup(XElement sectionGroup, string sectionGroupName,
            Application oneNoteApp, string notebookId, XmlNamespaceManager xnm, NoteLinkManager.AnalyzeDepth linkDepth, bool force, bool deleteNotes)
        {
            string sectionGroupId = (string)sectionGroup.Attribute("ID");

            if (!string.IsNullOrEmpty(sectionGroupName))
            {
                Logger.LogMessage("Обработка группы секций '{0}'", sectionGroupName);
                Logger.MoveLevel(1);
            }

            foreach (var subSectionGroup in sectionGroup.XPathSelectElements("one:SectionGroup", xnm))
            {                
                string subSectionGroupName = (string)subSectionGroup.Attribute("name");
                ProcessSectionGroup(subSectionGroup, subSectionGroupName, oneNoteApp, notebookId, xnm, linkDepth, force, deleteNotes);
            }

            foreach (var subSection in sectionGroup.XPathSelectElements("one:Section", xnm))
            {
                ProcessSection(subSection, sectionGroupId, oneNoteApp, notebookId, xnm, linkDepth, force, deleteNotes);
            }

            if (!string.IsNullOrEmpty(sectionGroupName))
            {
                Logger.MoveLevel(-1);
            }
        }

        private static void ProcessSection(XElement section, string sectionGroupId,
            Application oneNoteApp, string notebookId, XmlNamespaceManager xnm, NoteLinkManager.AnalyzeDepth linkDepth, bool force, bool deleteNotes)
        {
            string sectionId = (string)section.Attribute("ID");
            string sectionName = (string)section.Attribute("name");

            Logger.LogMessage("Обработка секции '{0}'", sectionName);
            Logger.MoveLevel(1);

            foreach (var page in section.XPathSelectElements("one:Page", xnm))
            {
                string pageId = (string)page.Attribute("ID");
                string pageName = (string)page.Attribute("name");

                Logger.LogMessage("Обработка страницы '{0}'", pageName);

                Logger.MoveLevel(1);
                if (deleteNotes)
                    NoteLinkManager.DeletePageNotes(oneNoteApp, notebookId, sectionGroupId, sectionId, pageId, pageName);
                else
                    NoteLinkManager.LinkPageVerses(oneNoteApp, notebookId, sectionGroupId, sectionId, pageId, linkDepth, force);
                Logger.MoveLevel(-1);
            }

            Logger.MoveLevel(-1);
        }
    }
}
