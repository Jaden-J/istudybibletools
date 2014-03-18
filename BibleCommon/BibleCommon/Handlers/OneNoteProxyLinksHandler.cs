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

namespace BibleCommon.Handlers
{
    public class OneNoteProxyLinksHandler : IProtocolHandler
    {
        private const string _protocolName = "bnONL:";

        public string ProtocolName
        {
            get { return _protocolName; }
        }

        public static string GetCommandUrlStatic(ref Application oneNoteApp, string pageId, string objectId, bool autoCommit = false)
        {
            var changed = false;
            var pageInfo = ApplicationCache.Instance.GetPageContent(ref oneNoteApp, pageId, ApplicationCache.PageType.NotePage);
            var bnPId = OneNoteUtils.GetElementMetaData(pageInfo.Content.Root, Consts.Constants.Key_PageId, pageInfo.Xnm);            

            if (string.IsNullOrEmpty(bnPId))
            {
                bnPId = Guid.NewGuid().ToString();
                OneNoteUtils.UpdateElementMetaData(pageInfo.Content.Root, Consts.Constants.Key_PageId, bnPId, pageInfo.Xnm);
                changed = true;
            }

            var bnOeId = string.Empty;
            var el = pageInfo.Content.XPathSelectElement(string.Format("//one:OE[@objectID=\"{0}\"]", objectId), pageInfo.Xnm);
            if (el != null)
            {   
                bnOeId = OneNoteUtils.GetElementMetaData(el, Consts.Constants.Key_OEId, pageInfo.Xnm);
                if (string.IsNullOrEmpty(bnOeId))
                {
                    bnOeId = Guid.NewGuid().ToString();
                    OneNoteUtils.UpdateElementMetaData(el, Consts.Constants.Key_OEId, bnOeId, pageInfo.Xnm);
                    changed = true;
                }                
            }

            if (changed && autoCommit)
                ApplicationCache.Instance.CommitModifiedPage(ref oneNoteApp, pageInfo, true);

            return string.Format("{0}{1};{2}", _protocolName, bnPId, bnOeId);
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

            var parts = args[0].Split(';');
            var bnPId = parts[0];
            var bnOeId = parts[1];

            var oneNoteApp = OneNoteUtils.CreateOneNoteAppSafe();

            try
            {
                var hierarchyInfo = ApplicationCache.Instance.GetHierarchy(ref oneNoteApp, null, HierarchyScope.hsPages);

                var pEl = hierarchyInfo.Content
                    .XPathSelectElement(string.Format("//one:Page[./one:Meta[@name=\"{0}\" and @content=\"{1}\"]]", Consts.Constants.Key_PageId, bnPId), hierarchyInfo.Xnm);
                var pId = (string)pEl.Attribute("ID");

                var pageInfo = ApplicationCache.Instance.GetPageContent(ref oneNoteApp, pId, ApplicationCache.PageType.NotePage);
                var oeId = string.Empty;
                if (!string.IsNullOrEmpty(bnOeId))
                {
                    var oeEl = pageInfo.Content
                        .XPathSelectElement(string.Format("//one:OE[./one:Meta[@name=\"{0}\" and @content=\"{1}\"]]", Consts.Constants.Key_OEId, bnOeId), pageInfo.Xnm);
                    oeId = (string)oeEl.Attribute("objectID");
                }

                oneNoteApp.NavigateTo(pId, oeId);

                OneNoteUtils.SetActiveCurrentWindow(ref oneNoteApp);

                return true;
            }
            finally
            {
                OneNoteUtils.ReleaseOneNoteApp(ref oneNoteApp);
            }
        }


        string IProtocolHandler.GetCommandUrl(string args)
        {
            return string.Format("{0}:{1}", ProtocolName, args);
        }
    }
}
