using BibleCommon.Contracts;
using BibleCommon.Helpers;
using BibleCommon.Services;
using Microsoft.Office.Interop.OneNote;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Xml;
using System.Xml.XPath;
using BibleCommon.Common;

namespace BibleCommon.Handlers
{
    public class OneNoteProxyLinksHandler : IProtocolHandler
    {
        private const string _protocolName = "bnONL:";

        public string ProtocolName
        {
            get { return _protocolName; }
        }

        private Dictionary<string, string> _notebookIds;
        private static object _locker = new object();

        /// <summary>
        /// 
        /// </summary>
        /// <param name="oneNoteApp"></param>
        /// <param name="notebookName">can be null or empty</param>
        /// <param name="pageId"></param>
        /// <param name="objectId"></param>
        /// <param name="autoCommit"></param>
        /// <returns></returns>
        public static string GetCommandUrlStatic(ref Application oneNoteApp, LinkId pageObjectInfo, bool autoCommit)
        {
            var changed = false;
            var pageInfo = ApplicationCache.Instance.GetPageContent(ref oneNoteApp, pageObjectInfo.PageId, ApplicationCache.PageType.NotePage);
            var bnPId = OneNoteUtils.GetElementMetaData(pageInfo.Content.Root, Consts.Constants.Key_PageId, pageInfo.Xnm);            

            if (string.IsNullOrEmpty(bnPId))
            {
                bnPId = Guid.NewGuid().ToString();
                OneNoteUtils.UpdateElementMetaData(pageInfo.Content.Root, Consts.Constants.Key_PageId, bnPId, pageInfo.Xnm);
                pageInfo.WasModified = true;
                changed = true;
            }

            var bnOeId = string.Empty;
            if (!string.IsNullOrEmpty(pageObjectInfo.ObjectId))
            {
                var el = pageInfo.Content.XPathSelectElement(string.Format("//one:OE[@objectID=\"{0}\"]", pageObjectInfo.ObjectId), pageInfo.Xnm);
                if (el != null)
                {
                    bnOeId = OneNoteUtils.GetElementMetaData(el, Consts.Constants.Key_OEId, pageInfo.Xnm);
                    if (string.IsNullOrEmpty(bnOeId))
                    {
                        bnOeId = Guid.NewGuid().ToString();
                        OneNoteUtils.UpdateElementMetaData(el, Consts.Constants.Key_OEId, bnOeId, pageInfo.Xnm);
                        pageInfo.WasModified = true;
                        changed = true;
                    }
                }
            }

            if (changed && autoCommit)
                ApplicationCache.Instance.CommitModifiedPage(ref oneNoteApp, pageInfo, true);

            return string.Format("{0}{1};{2};{3}", _protocolName, pageObjectInfo.NotebookName, bnPId, bnOeId);
        }

        public void ExecuteCommand(params string[] args)
        {
            try
            {
                TryExecuteCommand(args);
            }
            catch (Exception ex)
            {
                FormLogger.LogError(ex);
            }  
        }

        public bool IsProtocolCommand(params string[] args)
        {
            return args.Length > 0 && args[0].StartsWith(ProtocolName, StringComparison.OrdinalIgnoreCase);
        }

        internal bool TryExecuteCommand(params string[] args)
        {
            if (args.Length == 0)
                throw new ArgumentNullException("args");

            var parts = Uri.UnescapeDataString(args[0]
                            .Split(new char[] { ':' })[1])
                            .Split(new char[] { ';', '&' });
            var notebookName = parts[0];
            var bnPId = parts[1];
            var bnOeId = parts[2];

            var oneNoteApp = OneNoteUtils.CreateOneNoteAppSafe();

            try
            {
                var pageId = GetOneNotePageId(ref oneNoteApp, notebookName, bnPId);               

                if (!string.IsNullOrEmpty(pageId))
                {
                    var objectId = GetOneNoteObjectId(ref oneNoteApp, pageId, bnOeId);

                    oneNoteApp.NavigateTo(pageId, objectId);

                    OneNoteUtils.SetActiveCurrentWindow(ref oneNoteApp);

                    return true;
                }
                else                
                    throw new Exception(BibleCommon.Resources.Constants.PageNotFound);                                    
            }
            finally
            {
                OneNoteUtils.ReleaseOneNoteApp(ref oneNoteApp);
            }
        }

        private string GetOneNotePageId(ref Application oneNoteApp, string notebookName, string bnPId, bool searchInAllNotebooks = false)
        {
            string notebookId = null;

            if (!searchInAllNotebooks && !string.IsNullOrEmpty(notebookName))
            {
                if (_notebookIds == null)
                {
                    lock (_locker)
                    {
                        if (_notebookIds == null)
                            _notebookIds = LoadNotebookIds(ref oneNoteApp);
                    }
                }

                if (_notebookIds.ContainsKey(notebookName))
                    notebookId = _notebookIds[notebookName];
            }

            var hierarchyInfo = ApplicationCache.Instance.GetHierarchy(ref oneNoteApp, notebookId, HierarchyScope.hsPages);

            var pEl = hierarchyInfo.Content
                .XPathSelectElement(string.Format("//one:Page[./one:Meta[@name=\"{0}\" and @content=\"{1}\"]]", Consts.Constants.Key_PageId, bnPId), hierarchyInfo.Xnm);

            if (pEl == null)
            {
                hierarchyInfo = ApplicationCache.Instance.GetHierarchy(ref oneNoteApp, notebookId, HierarchyScope.hsPages, true);

                pEl = hierarchyInfo.Content
                    .XPathSelectElement(string.Format("//one:Page[./one:Meta[@name=\"{0}\" and @content=\"{1}\"]]", Consts.Constants.Key_PageId, bnPId), hierarchyInfo.Xnm);
            }

            if (pEl != null)
                return (string)pEl.Attribute("ID");
            else if (!searchInAllNotebooks && !string.IsNullOrEmpty(notebookName))
                return GetOneNotePageId(ref oneNoteApp, notebookName, bnPId, true);
            else 
                return null;
        }

        private static Dictionary<string, string> LoadNotebookIds(ref Application oneNoteApp)
        {
            var notebooksDoc = ApplicationCache.Instance.GetHierarchy(ref oneNoteApp, null, HierarchyScope.hsNotebooks);
            var notebookIds = new Dictionary<string, string>();

            foreach (var notebookEl in notebooksDoc.Content.Root.XPathSelectElements("one:Notebook", notebooksDoc.Xnm))
            {
                var name = (string)notebookEl.Attribute("name");
                var id = (string)notebookEl.Attribute("ID");

                if (!string.IsNullOrEmpty(name) && !string.IsNullOrEmpty(id) && !notebookIds.ContainsKey(name))
                    notebookIds.Add(name, id);
            }

            return notebookIds;
        }

        private string GetOneNoteObjectId(ref Application oneNoteApp, string pageId, string bnOeId)
        {
            var pageInfo = ApplicationCache.Instance.GetPageContent(ref oneNoteApp, pageId, ApplicationCache.PageType.NotePage);
            
            if (!string.IsNullOrEmpty(bnOeId))
            {
                var oeEl = pageInfo.Content
                            .XPathSelectElement(string.Format("//one:OE[./one:Meta[@name=\"{0}\" and @content=\"{1}\"]]", Consts.Constants.Key_OEId, bnOeId), pageInfo.Xnm);

                if (oeEl == null)
                {
                    pageInfo = ApplicationCache.Instance.GetPageContent(ref oneNoteApp, pageId, ApplicationCache.PageType.NotePage, true, PageInfo.piBasic, false);
                    oeEl = pageInfo.Content
                            .XPathSelectElement(string.Format("//one:OE[./one:Meta[@name=\"{0}\" and @content=\"{1}\"]]", Consts.Constants.Key_OEId, bnOeId), pageInfo.Xnm);
                }

                if (oeEl != null)
                    return (string)oeEl.Attribute("objectID");
            }

            return null;
        }      


        string IProtocolHandler.GetCommandUrl(string args)
        {
            return string.Format("{0}:{1}", ProtocolName, args);
        }
    }
}
