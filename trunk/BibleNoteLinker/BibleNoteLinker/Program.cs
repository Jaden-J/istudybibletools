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
using BibleCommon.Consts;

namespace BibleNoteLinker
{
    public class Program
    {
        const string Arg_AllPages = "-allpages";
        const string Arg_DeleteNotes = "-deletenotes";
        const string Arg_Force = "-force";
        const string Arg_Changed = "-changed";
            
        public class Args
        {
            public bool AnalyzeAllPages { get; set; }  // false - значит только текущую
            public NoteLinkManager.AnalyzeDepth AnalyzeDepth { get; set; }
            public bool Force { get; set; }                        
            public bool LastChanged { get; set; }

            public Args()
            {
                this.AnalyzeAllPages = false;
                this.AnalyzeDepth = NoteLinkManager.AnalyzeDepth.Full;
                this.Force = false;                
                this.LastChanged = false;
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
                else if (argLower == Arg_Changed)
                    result.LastChanged = true;
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
            Application oneNoteApp = new Application();                    

            if (!SettingsManager.Instance.IsConfigured(oneNoteApp))
            {
                Logger.LogError(Constants.Error_SystemIsNotConfigures);
            }
            else
            {
                try
                {
                    Logger.LogMessage("Время старта: {0}", dtStart.ToLongTimeString());

                    Args userArgs = GetUserArgs(args);
                    
                    Logger.LogMessage("Уровень текущего анализа: '{0} ({1})'", userArgs.AnalyzeDepth, (int)userArgs.AnalyzeDepth);
                    if (userArgs.Force)
                        Logger.LogMessage("Анализируем ссылки в том числе");
                    if (userArgs.LastChanged)
                        Logger.LogMessage("Анализируем только последние модифицированные страницы");
                    

                    if (userArgs.AnalyzeAllPages)
                    {
                        if (!userArgs.LastChanged)
                            Logger.LogMessage("Старт обработки всех страниц");

                        ProcessNotebook(oneNoteApp, SettingsManager.Instance.NotebookId_BibleNotesPages, SettingsManager.Instance.SectionGroupId_BibleStudy, userArgs);
                        ProcessNotebook(oneNoteApp, SettingsManager.Instance.NotebookId_BibleComments, SettingsManager.Instance.SectionGroupId_BibleComments, userArgs);                        
                        ProcessNotebook(oneNoteApp, SettingsManager.Instance.NotebookId_BibleStudy, SettingsManager.Instance.SectionGroupId_BibleStudy, userArgs);                        
                    }
                    else
                    {
                        Logger.LogMessage("Старт обработки текущей страницы");

                        if (oneNoteApp.Windows.CurrentWindow != null)
                        {
                            string currentPageId = oneNoteApp.Windows.CurrentWindow.CurrentPageId;
                            if (!string.IsNullOrEmpty(currentPageId))
                            {
                                string currentSectionId = oneNoteApp.Windows.CurrentWindow.CurrentSectionId;
                                string currentSectionGroupId = oneNoteApp.Windows.CurrentWindow.CurrentSectionGroupId;
                                string currentNotebookId = oneNoteApp.Windows.CurrentWindow.CurrentNotebookId;

                                if (!string.IsNullOrEmpty(currentPageId))
                                {
                                    if (currentNotebookId == SettingsManager.Instance.NotebookId_BibleComments
                                        || currentNotebookId == SettingsManager.Instance.NotebookId_BibleStudy
                                        || currentNotebookId == SettingsManager.Instance.NotebookId_BibleNotesPages)
                                    {
                                        new NoteLinkManager(oneNoteApp).LinkPageVerses(currentSectionGroupId, currentSectionId, currentPageId, userArgs.AnalyzeDepth, userArgs.Force);
                                    }
                                    else
                                        Logger.LogError(string.Format("Текущая записная книжка не настроена на обработку программой {0}", Constants.ToolsName));
                                }
                                else
                                    Logger.LogError("Не найдено открытой страницы заметок");
                            }
                            else
                                Logger.LogError("Не найдено открытой страницы заметок");
                        }
                        else
                            Logger.LogError("Не найдено открытой записной книжки");
                    }                    

                    Logger.LogMessage("Обновление страниц 'Сводные заметок' в OneNote", true, false);          
                    OneNoteProxy.Instance.CommitAllModifiedPages(oneNoteApp, 
                        pageContent => pageContent.PageType == OneNoteProxy.PageType.NotesPage,
                        pagesCount => Logger.LogMessage(string.Format(" ({0})", GetRightPagesString(pagesCount)), false, false, false),
                        pageContent => Logger.LogMessage(".", false, false, false));
                    Logger.LogMessage(string.Empty, false, true, false);

                    Logger.LogMessage(string.Format("Обновление ссылок на страницы 'Сводные заметок' ({0})",
                        GetRightPagesString(OneNoteProxy.Instance.ProcessedBiblePages.Values.Count)), true, false);
                    var relinkNotesManager = new RelinkAllBibleNotesManager(oneNoteApp);
                    foreach (OneNoteProxy.BiblePageId processedBiblePageId in OneNoteProxy.Instance.ProcessedBiblePages.Values)
                    {
                        relinkNotesManager.RelinkBiblePageNotes(processedBiblePageId.SectionId, processedBiblePageId.PageId, 
                            processedBiblePageId.PageName, processedBiblePageId.ChapterPointer);
                        Logger.LogMessage(".", false, false, false);
                    }
                    Logger.LogMessage(string.Empty, false, true, false);

                    Logger.LogMessage("Обновление страниц в OneNote", true, false);
                    OneNoteProxy.Instance.CommitAllModifiedPages(oneNoteApp,
                        null,
                        pagesCount => Logger.LogMessage(string.Format(" ({0})", GetRightPagesString(pagesCount)), false, false, false),
                        pageContent => Logger.LogMessage(".", false, false, false));
                    Logger.LogMessage(string.Empty, false, true, false);

                    //Сортировка страниц 'Сводные заметок'
                    foreach (var sortPageInfo in OneNoteProxy.Instance.SortVerseLinkPagesInfo)
                    {
                        VerseLinkManager.SortVerseLinkPages(oneNoteApp, 
                            sortPageInfo.SectionId, sortPageInfo.PageId, sortPageInfo.ParentPageId, sortPageInfo.PageLevel);                    
                    }                    

                    Logger.LogMessage("Обновление иерархии в OneNote", true, false);
                    OneNoteProxy.Instance.CommitAllModifiedHierarchy(oneNoteApp,                        
                        pagesCount => Logger.LogMessage(string.Format(" ({0})", GetRightPagesString(pagesCount)), false, false, false),
                        pageContent => Logger.LogMessage(".", false, false, false));
                    Logger.LogMessage(string.Empty, false, true, false);
                }
                catch (Exception ex)
                {
                    Logger.LogError(ex);
                }
            }

            Logger.LogMessage("Времени затрачено: {0}", DateTime.Now.Subtract(dtStart));
            

            if (Logger.ErrorWasLogged)            
                Console.WriteLine("Во время работы программы произошли ошибки");                            
            else
                Logger.LogMessage("Успешно завершено");

            Logger.Done();

            if (Logger.ErrorWasLogged)
                Console.ReadKey();
        }

        private static string GetRightPagesString(int pagesCount)
        {
            string s = "страниц";
            int tempPagesCount = pagesCount;

            tempPagesCount = tempPagesCount % 100;
            if (!(tempPagesCount >= 10 && tempPagesCount <= 20))
            {
                tempPagesCount = tempPagesCount % 10;

                if (tempPagesCount == 1)
                    s = "страница";
                else if (tempPagesCount >= 2 && tempPagesCount <= 4)
                    s = "страницы";
            }

            return string.Format("{0} {1}", pagesCount, s);
        }

        private static void ProcessNotebook(Application oneNoteApp, string notebookId, string sectionGroupId, Args userArgs)
        {
            Logger.LogMessage("Обработка записной книжки: '{0}'", OneNoteUtils.GetHierarchyElementName(oneNoteApp, notebookId));  // чтобы точно убедиться

            OneNoteProxy.HierarchyElement notebookPages = OneNoteProxy.Instance.GetHierarchy(oneNoteApp, notebookId, HierarchyScope.hsPages);

            Logger.MoveLevel(1);
            ProcessRootSectionGroup(oneNoteApp, notebookId, notebookPages.Content, sectionGroupId, notebookPages.Xnm, userArgs.AnalyzeDepth, userArgs.Force, userArgs.LastChanged);
            Logger.MoveLevel(-1);
        }

        private static void ProcessLastChangedPages(Application oneNoteApp,
            XElement rootSectionGroup, XmlNamespaceManager xnm, NoteLinkManager.AnalyzeDepth linkDepth, bool force)
        {
            foreach (XElement page in rootSectionGroup.XPathSelectElements(".//one:Page", xnm))
            {
                XAttribute lastModifiedDateAttribute = page.Attribute("lastModifiedTime");
                if (lastModifiedDateAttribute != null)
                {
                    DateTime lastModifiedDate = DateTime.Parse(lastModifiedDateAttribute.Value);

                    bool needToAnalyze = true;

                    string lastAnalyzeTime = OneNoteUtils.GetPageMetaData(oneNoteApp, page, Constants.Key_LatestAnalyzeTime, xnm);
                    if (!string.IsNullOrEmpty(lastAnalyzeTime) && lastModifiedDate <= DateTime.Parse(lastAnalyzeTime).ToLocalTime())
                        needToAnalyze = false;

                    if (needToAnalyze)
                    {
                        string sectionGroupId = string.Empty;
                        XElement sectionGroup = page.Parent.Parent;
                        if (sectionGroup.Name.LocalName == "SectionGroup" && !OneNoteUtils.IsRecycleBin(sectionGroup))
                            sectionGroupId = (string)sectionGroup.Attribute("ID").Value;

                        string sectionId = (string)page.Parent.Attribute("ID").Value;

                        ProcessPage(oneNoteApp, page, sectionGroupId, sectionId, linkDepth, force);
                    }
                }
            }
        }

        private static void ProcessRootSectionGroup(Application oneNoteApp, string notebookId, XDocument doc, string sectionGroupId,
            XmlNamespaceManager xnm, NoteLinkManager.AnalyzeDepth linkDepth, bool force, bool lastChanged)
        {
            XElement sectionGroup = string.IsNullOrEmpty(sectionGroupId) 
                                        ? doc.Root 
                                        : doc.Root.XPathSelectElement(
                                                string.Format("one:SectionGroup[@ID='{0}']", sectionGroupId), xnm);

            if (sectionGroup != null)
            {
                if (lastChanged)
                    ProcessLastChangedPages(oneNoteApp, sectionGroup, xnm, linkDepth, force);
                else
                    ProcessSectionGroup(sectionGroup, sectionGroupId, oneNoteApp, notebookId, xnm, linkDepth, force);
            }
            else
                Logger.LogError("Не удаётся найти группу секций '{0}'", sectionGroupId);
        }

        private static void ProcessSectionGroup(XElement sectionGroup, string sectionGroupId,
            Application oneNoteApp, string notebookId, XmlNamespaceManager xnm, NoteLinkManager.AnalyzeDepth linkDepth, bool force)
        {
            string sectionGroupName = (string)sectionGroup.Attribute("name");

            if (OneNoteUtils.IsRecycleBin(sectionGroup))
                return;

            if (!string.IsNullOrEmpty(sectionGroupName))
            {
                Logger.LogMessage("Обработка группы секций '{0}'", sectionGroupName);
                Logger.MoveLevel(1);
            }

            foreach (var subSectionGroup in sectionGroup.XPathSelectElements("one:SectionGroup", xnm).Where(sg => !OneNoteUtils.IsRecycleBin(sg)))
            {                
                string subSectionGroupId = (string)subSectionGroup.Attribute("ID");
                ProcessSectionGroup(subSectionGroup, subSectionGroupId, oneNoteApp, notebookId, xnm, linkDepth, force);
            }

            foreach (var subSection in sectionGroup.XPathSelectElements("one:Section", xnm))
            {
                ProcessSection(subSection, sectionGroupId, oneNoteApp, notebookId, xnm, linkDepth, force);
            }

            if (!string.IsNullOrEmpty(sectionGroupName))
            {
                Logger.MoveLevel(-1);
            }
        }

        private static void ProcessSection(XElement section, string sectionGroupId,
            Application oneNoteApp, string notebookId, XmlNamespaceManager xnm, NoteLinkManager.AnalyzeDepth linkDepth, bool force)
        {
            string sectionId = (string)section.Attribute("ID");
            string sectionName = (string)section.Attribute("name");

            Logger.LogMessage("Обработка секции '{0}'", sectionName);
            Logger.MoveLevel(1);

            foreach (var page in section.XPathSelectElements("one:Page", xnm))
            {
                ProcessPage(oneNoteApp, page, sectionGroupId, sectionId, linkDepth, force);
            }

            Logger.MoveLevel(-1);
        }

        private static void ProcessPage(Application oneNoteApp, XElement page, string sectionGroupId, string sectionId,
            NoteLinkManager.AnalyzeDepth linkDepth, bool force)
        {
            string pageId = (string)page.Attribute("ID");
            string pageName = (string)page.Attribute("name");

            Logger.LogMessage("Обработка страницы '{0}'", pageName);

            Logger.MoveLevel(1);
            
            new NoteLinkManager(oneNoteApp).LinkPageVerses(sectionGroupId, sectionId, pageId, linkDepth, force);

            Logger.MoveLevel(-1);
        }
    }
}
