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
                    
                    string currentNotebookId = GetBibleNotebookId(oneNoteApp, SettingsManager.Instance.NotebookName);

                    if (!string.IsNullOrEmpty(currentNotebookId))
                    {
                        Logger.LogMessage("Имя записной книжки: {0}", Utils.GetHierarchyElementName(oneNoteApp, currentNotebookId));  // чтобы точно убедиться

                        string hierarchyXml;
                        oneNoteApp.GetHierarchy(currentNotebookId, HierarchyScope.hsPages, out hierarchyXml);
                        XmlNamespaceManager xnm;
                        XDocument notebookDoc = Utils.GetXDocument(hierarchyXml, out xnm);

                        if (userArgs.DeleteNotes)
                            ProcessRootSectionGroup(oneNoteApp, currentNotebookId, notebookDoc, SettingsManager.Instance.BibleSectionGroupName, xnm, userArgs.AnalyzeDepth, userArgs.Force, userArgs.DeleteNotes);
                        else
                        {
                            ProcessRootSectionGroup(oneNoteApp, currentNotebookId, notebookDoc, SettingsManager.Instance.StudyBibleSectionGroupName, xnm, userArgs.AnalyzeDepth, userArgs.Force, userArgs.DeleteNotes);
                            ProcessRootSectionGroup(oneNoteApp, currentNotebookId, notebookDoc, SettingsManager.Instance.ResearchSectionGroupName, xnm, userArgs.AnalyzeDepth, userArgs.Force, userArgs.DeleteNotes);
                        }
                    }
                    else
                    {
                        Logger.LogError(string.Format("Не найдено записной книжки '{0}'.", SettingsManager.Instance.NotebookName));
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
                                NoteLinkManager.DeletePageNotes(oneNoteApp, currentNotebookId, currentSectionGroupId, currentSectionId, currentPageId, Utils.GetHierarchyElementName(oneNoteApp,currentPageId));
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
      
        private static string GetBibleNotebookId(Application oneNoteApp, string notebookName)
        {
            string xml;
            XmlNamespaceManager xnm;
            oneNoteApp.GetHierarchy(null, HierarchyScope.hsNotebooks, out xml);
            XDocument doc = Utils.GetXDocument(xml, out xnm);            
            XElement bibleNotebook = doc.Root.XPathSelectElement(string.Format("one:Notebook[@name='{0}']", notebookName), xnm);
            if (bibleNotebook != null)
            {
                return (string)bibleNotebook.Attribute("ID");                
            }

            return null;
        }

        private static void ProcessRootSectionGroup(Application oneNoteApp, string notebookId, XDocument doc, string sectionGroupName,
            XmlNamespaceManager xnm, NoteLinkManager.AnalyzeDepth linkDepth, bool force, bool deleteNotes)
        {
            XElement sectionGroup = doc.Root.XPathSelectElement(
                            string.Format("one:SectionGroup[@name='{0}']", sectionGroupName), xnm);                        

            if (sectionGroup != null)
                ProcessSectionGroup(sectionGroup, oneNoteApp, notebookId, xnm, linkDepth, force, deleteNotes);
            else
                Logger.LogError("Не удаётся найти группу секций '{0}'", sectionGroupName);
        }

        private static void ProcessSectionGroup(XElement sectionGroup,
            Application oneNoteApp, string notebookId, XmlNamespaceManager xnm, NoteLinkManager.AnalyzeDepth linkDepth, bool force, bool deleteNotes)
        {
            string sectionGroupId = (string)sectionGroup.Attribute("ID");
            string sectionGroupName = (string)sectionGroup.Attribute("name");

            Logger.LogMessage("Обработка группы секций '{0}'", sectionGroupName);
            Logger.MoveLevel(1);

            foreach (var subSectionGroup in sectionGroup.XPathSelectElements("one:SectionGroup", xnm))
            {
                ProcessSectionGroup(subSectionGroup, oneNoteApp, notebookId, xnm, linkDepth, force, deleteNotes);
            }

            foreach (var subSection in sectionGroup.XPathSelectElements("one:Section", xnm))
            {
                ProcessSection(subSection, sectionGroupId, oneNoteApp, notebookId, xnm, linkDepth, force, deleteNotes);
            }

            Logger.MoveLevel(-1);
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
