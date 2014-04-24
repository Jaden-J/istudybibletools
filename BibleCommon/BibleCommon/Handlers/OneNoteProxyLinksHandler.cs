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
using BibleCommon.Consts;

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
        private static readonly object _locker = new object();

        /// <summary>
        /// 
        /// </summary>
        /// <param name="oneNoteApp"></param>
        /// <param name="notebookName">can be null or empty</param>
        /// <param name="pageId"></param>
        /// <param name="objectId"></param>
        /// <param name="autoCommit"></param>
        /// <returns></returns>
        public static string GetCommandUrlStatic(ref Application oneNoteApp, LinkId pageObjectInfo, bool autoCommit, bool checkForDuplicateId)
        {
            var changed = false;

            var bnPId = pageObjectInfo.IdType == IdType.Custom ? pageObjectInfo.PageId : string.Empty;
            var bnOeId = pageObjectInfo.IdType == IdType.Custom ? pageObjectInfo.ObjectId : string.Empty;

            if (pageObjectInfo.IdType == IdType.OneNote)
            {
                var pageInfo = ApplicationCache.Instance.GetPageContent(ref oneNoteApp, pageObjectInfo.PageId, ApplicationCache.PageType.NotePage);
                var updatePageResult = GetOrUpdateBnPId(ref oneNoteApp, pageInfo.Content.Root, checkForDuplicateId, pageInfo.Xnm);
                bnPId = updatePageResult.Id;
                if (updatePageResult.Changed)
                {
                    pageInfo.WasModified = true;
                    changed = true;
                }

                if (!string.IsNullOrEmpty(pageObjectInfo.ObjectId))
                {
                    var el = pageInfo.Content.XPathSelectElement(string.Format("//one:OE[@objectID=\"{0}\"]", pageObjectInfo.ObjectId), pageInfo.Xnm);
                    var updateElResult = GetOrUpdateBnOeId(el, pageInfo.Xnm);
                    bnOeId = updateElResult.Id;
                    if (updateElResult.Changed)
                    {
                        pageInfo.WasModified = true;
                        changed = true;
                    }
                }

                if (changed && autoCommit)
                    ApplicationCache.Instance.CommitModifiedPage(ref oneNoteApp, pageInfo, true);
            }

            return NavigateToHandler.AddIdParametersToLink(
                        string.Format("{0}{1};{2};{3}", _protocolName, pageObjectInfo.NotebookName, bnPId, bnOeId),
                        pageObjectInfo.PageId, pageObjectInfo.ObjectId);
        }

        public class UpdateResult
        {
            public string Id { get; set;}
            public bool Changed { get; set; }
        }

        public static UpdateResult GetOrUpdateBnPId(ref Application oneNoteApp, XElement pageEl, bool checkForDuplicateId, XmlNamespaceManager xnm)
        {
            var result = new UpdateResult();
            result.Id = OneNoteUtils.GetElementMetaData(pageEl, Consts.Constants.Key_PId, xnm);

            if (checkForDuplicateId && !string.IsNullOrEmpty(result.Id))
            {
                if (ApplicationCache.Instance.PagesCustomIds.ContainsKey(result.Id))  // почему не нужно всегда проверять, а только при checkForDuplicateId? Потому что вот эта операция (собирание PagesCustomIds) очень дорогая! 
                {
                    var oneNoteId = (string)pageEl.Attribute("ID");
                    if (ApplicationCache.Instance.PagesCustomIds[result.Id] != oneNoteId)  // то есть уже есть другая такая страница (с таким же CustomId)
                    {
                        result.Id = null;   // пока так, так как сейчас нет возможности проследить, какая страница была создана первой. Так как копируется и значение атрибута "dateTime". В 4.0 версии можно будет в базе хранить дату создания страницы (или дату первого анализа). И уже по этому параметру понимать, у какой страницы мы оставляем этот CustomId, а какой генерируем новый.
                    }
                }
            }

            if (string.IsNullOrEmpty(result.Id))
            {
                result.Id = GeneratePId();
                OneNoteUtils.UpdateElementMetaData(pageEl, Consts.Constants.Key_PId, result.Id, xnm);
                result.Changed = true;
            }            

            return result;
        }

        public static string GeneratePId()
        {
            return Guid.NewGuid().ToString();
        }

        public static UpdateResult GetOrUpdateBnOeId(XElement el, XmlNamespaceManager xnm)
        {
            var result = new UpdateResult();
            if (el != null)
            {
                result.Id = OneNoteUtils.GetElementMetaData(el, Consts.Constants.Key_OEId, xnm);
                if (string.IsNullOrEmpty(result.Id))
                {
                    result.Id = Guid.NewGuid().ToString();
                    OneNoteUtils.UpdateElementMetaData(el, Consts.Constants.Key_OEId, result.Id, xnm);                    
                    result.Changed = true;
                }
            }
            return result;
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

            var oneNoteApp = OneNoteUtils.CreateOneNoteAppSafe();

            try
            {
                if (!NavigateToHandler.TryToRedirectByIds(ref oneNoteApp, args[0]))
                {
                    if (!TryToRedirectByMetadataId(ref oneNoteApp, args[0]))
                    {
                        return false;
                    }
                }

                return true;
            }
            finally
            {
                OneNoteUtils.ReleaseOneNoteApp(ref oneNoteApp);
            }
        }

        private bool TryToRedirectByMetadataId(ref Application oneNoteApp, string args)
        {
            var parts = Uri.UnescapeDataString(args
                               .Split(new char[] { ':' })[1])
                               .Split(new char[] { ';', '&' });
            var notebookName = parts[0];
            var bnPId = parts[1];
            var bnOeId = parts[2];

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
                .XPathSelectElement(string.Format("//one:Page[./one:Meta[@name=\"{0}\" and @content=\"{1}\"]]", Consts.Constants.Key_PId, bnPId), hierarchyInfo.Xnm);

            if (pEl == null)
            {
                hierarchyInfo = ApplicationCache.Instance.GetHierarchy(ref oneNoteApp, notebookId, HierarchyScope.hsPages, true);

                pEl = hierarchyInfo.Content
                    .XPathSelectElement(string.Format("//one:Page[./one:Meta[@name=\"{0}\" and @content=\"{1}\"]]", Consts.Constants.Key_PId, bnPId), hierarchyInfo.Xnm);
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
