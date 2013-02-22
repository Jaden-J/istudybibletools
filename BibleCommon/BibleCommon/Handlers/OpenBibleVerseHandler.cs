using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BibleCommon.Contracts;
using BibleCommon.Common;
using BibleCommon.Helpers;
using System.Diagnostics;
using BibleCommon.Services;
using System.Xml.XPath;
using System.Xml;
using Microsoft.Office.Interop.OneNote;
using System.Runtime.InteropServices;

namespace BibleCommon.Handlers
{
    public class OpenBibleVerseHandler : IProtocolHandler
    {
        private const string _protocolName = "isbtBibleVerse:";

        public string ProtocolName
        {
            get { return _protocolName; }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="vp"></param>
        /// <param name="moduleName">может быть null</param>
        /// <returns></returns>
        public string GetCommandUrl(VersePointer vp, string moduleName)
        {
            return GetCommandUrlStatic(vp, moduleName);
        }

        public static string GetCommandUrlStatic(VersePointer vp, string moduleName)
        {
            return string.Format("{0}{1}/{2} {3};{4}", _protocolName, moduleName, vp.Book.Index, vp.VerseNumber, vp.OriginalVerseName);
        }

        public bool IsProtocolCommand(params string[] args)
        {
            return args.Length > 0 && args[0].StartsWith(ProtocolName, StringComparison.OrdinalIgnoreCase);
        }

        public void ExecuteCommand(params string[] args)
        {
            Application oneNoteApp = null;
            try
            {
                var index = args[0].IndexOf(";");
                if (index == -1)
                    throw new ArgumentException(string.Format("Ivalid versePointer args: {0}", args[0]));

                oneNoteApp = new Application();

                if (!SettingsManager.Instance.IsConfigured(ref oneNoteApp))                
                    throw new NotConfiguredException();                                    

                var verseString = Uri.UnescapeDataString(args[0].Substring(index + 1));

                var vp = new VersePointer(verseString);

                if (vp.IsValid)                
                    GoToVerse(ref oneNoteApp, vp);                
                else
                    throw new Exception(BibleCommon.Resources.Constants.BibleVersePointerCanNotParseString);
            }
            catch (Exception ex)
            {
                FormLogger.LogError(ex);
            }
            finally
            {
                if (oneNoteApp != null)
                {
                    Marshal.ReleaseComObject(oneNoteApp);
                    oneNoteApp = null;
                }
            }
        }

        private bool GoToVerse(ref Application oneNoteApp, VersePointer vp)
        {
            var result = HierarchySearchManager.GetHierarchyObject(ref oneNoteApp, SettingsManager.Instance.NotebookId_Bible, vp, HierarchySearchManager.FindVerseLevel.OnlyFirstVerse);

            if (result.ResultType != HierarchySearchManager.HierarchySearchResultType.NotFound
                && (result.HierarchyStage == HierarchySearchManager.HierarchyStage.ContentPlaceholder || result.HierarchyStage == HierarchySearchManager.HierarchyStage.Page))
            {
                string hierarchyObjectId = !string.IsNullOrEmpty(result.HierarchyObjectInfo.PageId)
                    ? result.HierarchyObjectInfo.PageId : result.HierarchyObjectInfo.SectionId;

                NavigateTo(ref oneNoteApp, hierarchyObjectId, result.HierarchyObjectInfo.GetAllObjectsIds().ToArray());
                return true;
            }
            else
                Logger.LogError(BibleCommon.Resources.Constants.BibleVersePointerCanNotFindPlace);

            return false;
        }

        private void NavigateTo(ref Application oneNoteApp, string pageId, params HierarchySearchManager.VerseObjectInfo[] objectsIds)
        {
            if (objectsIds.Length > 0 && !string.IsNullOrEmpty(objectsIds[0].ObjectHref))
            {
                var linksHandler = new NavigateToHandler();
                if (linksHandler.IsProtocolCommand(objectsIds[0].ObjectHref))
                    linksHandler.ExecuteCommand(objectsIds[0].ObjectHref);
                else
                    Process.Start(objectsIds[0].ObjectHref);   // иначе, если делать через NavigateTo, то когда, например, дропбокс изменит имя файла секции (сделает маленькими буквами) - ID меняется и выдаётся ошибка.
            }
            else
            {
                OneNoteUtils.UseOneNoteAPI(ref oneNoteApp, (oneNoteAppSafe) =>
                {
                    oneNoteAppSafe.NavigateTo(pageId, objectsIds.Length > 0 ? objectsIds[0].ObjectId : null);
                });
            }

            if (objectsIds.Length > 1)
            {
                XmlNamespaceManager xnm;
                var pageDoc = OneNoteUtils.GetPageContent(ref oneNoteApp, pageId, PageInfo.piSelection, out xnm);
                OneNoteLocker.UnlockCurrentSection(ref oneNoteApp);

                foreach (var objectId in objectsIds.Skip(1))
                {
                    var el = pageDoc.Root.XPathSelectElement(string.Format("//one:OE[@objectID=\"{0}\"]/one:T", objectId), xnm);
                    if (el != null)
                        el.SetAttributeValue("selected", "all");
                }

                OneNoteUtils.UpdatePageContentSafe(ref oneNoteApp, pageDoc, xnm);
            }
        }

        string IProtocolHandler.GetCommandUrl(string args)
        {
            return string.Format("{0}:{1}", ProtocolName, args);
        }      
    }
}
