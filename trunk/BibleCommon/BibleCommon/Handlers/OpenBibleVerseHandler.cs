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
            return string.Format("{0}{1}/{2} {3};{4}", 
                _protocolName, 
                moduleName, 
                vp.Book.Index,                 
                vp.GetFullMultiVerseString(),                
                vp.GetFriendlyFullVerseName());
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
                var parts = args[0].Split(new char[] { ';', '&' });
                if (parts.Length < 2)
                    throw new ArgumentException(string.Format("Ivalid versePointer args: {0}", args[0]));

                oneNoteApp = OneNoteUtils.CreateOneNoteAppSafe();

                var verseString = Uri.UnescapeDataString(parts[1]);

                var vp = new VersePointer(verseString);

                if (vp.IsValid)
                    GoToVerse(ref oneNoteApp, vp);
                else
                    throw new Exception(BibleCommon.Resources.Constants.BibleVersePointerCanNotParseString);
            }
            catch (InvalidModuleException imEx)
            {
                FormLogger.LogError(BibleCommon.Resources.Constants.Error_SystemIsNotConfigured + Environment.NewLine + imEx.Message);
            }
            catch (Exception ex)
            {
                FormLogger.LogError(ex);
            }
            finally
            {
                OneNoteUtils.ReleaseOneNoteApp(ref oneNoteApp);
            }
        }

        private bool GoToVerse(ref Application oneNoteApp, VersePointer vp)
        {
            var result = HierarchySearchManager.GetHierarchyObject(ref oneNoteApp, SettingsManager.Instance.NotebookId_Bible, ref vp, HierarchySearchManager.FindVerseLevel.OnlyFirstVerse);

            if (result.ResultType != BibleHierarchySearchResultType.NotFound
                && (result.HierarchyStage == BibleHierarchyStage.ContentPlaceholder || result.HierarchyStage == BibleHierarchyStage.Page))
            {
                string hierarchyObjectId = !string.IsNullOrEmpty(result.HierarchyObjectInfo.PageId)
                    ? result.HierarchyObjectInfo.PageId : result.HierarchyObjectInfo.SectionId;

                NavigateTo(ref oneNoteApp, vp, hierarchyObjectId, result.HierarchyObjectInfo.GetAllObjectsIds().ToArray());
                return true;
            }
            else
                Logger.LogError(BibleCommon.Resources.Constants.BibleVersePointerCanNotFindPlace);

            return false;
        }

        private void NavigateTo(ref Application oneNoteApp, VersePointer vp, string pageId, params VerseObjectInfo[] objectsIds)
        {
            var toCleanCacheAndRetry = false;
            if (!NavigateToHandler.TryToRedirectByIds(ref oneNoteApp, pageId, objectsIds.Length > 0 ? objectsIds[0].ObjectId : null))
            {
                if (objectsIds.Length > 0)
                {
                    var linkHref = objectsIds[0].Href;
                    if (!string.IsNullOrEmpty(linkHref))
                    {
                        var linksHandler = new NavigateToHandler();
                        if (linksHandler.IsProtocolCommand(linkHref))
                        {
                            if (!linksHandler.TryExecuteCommand(linkHref))
                                toCleanCacheAndRetry = true;
                        }
                        else
                        {
                            if (!NavigateToHandler.TryToRedirectByUrl(ref oneNoteApp, linkHref))
                                toCleanCacheAndRetry = true;
                        }
                    }
                    else
                        toCleanCacheAndRetry = true;                    
                }
                else
                    toCleanCacheAndRetry = true;                    
            }

            if (toCleanCacheAndRetry)
            {
                if (OneNoteProxy.Instance.IsBibleVersesLinksCacheActive)
                {
                    OneNoteProxy.Instance.CleanBibleVersesLinksCache(true);
                    Logger.LogWarning(BibleCommon.Resources.Constants.BibleVersesLinksCacheWasCleaned);                    

                    GoToVerse(ref oneNoteApp, vp);
                    return;
                }
            }

            OneNoteUtils.SetActiveCurrentWindow(ref oneNoteApp);

            //Пока не используем такую возможность, потому что Библия по умоланию заблокирована
            //SelectOtherVerses(ref oneNoteApp, pageId, objectsIds);
        }

        private void SelectOtherVerses(ref Application oneNoteApp, string pageId, params VerseObjectInfo[] objectsIds)
        {
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
