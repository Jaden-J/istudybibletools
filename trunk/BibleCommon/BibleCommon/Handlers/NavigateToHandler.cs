using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BibleCommon.Contracts;
using BibleCommon.Consts;
using BibleCommon.Services;
using System.Runtime.InteropServices;
using System.IO;
using System.Threading;
using BibleCommon.Helpers;

namespace BibleCommon.Handlers
{
    public class NavigateToHandler : IProtocolHandler
    {
        private const string _protocolName = "isbtOpen:";
        private const int NavigateAttemptsCount = 10;

        private static string ProtocolFullString
        {
            get
            {
                return string.Format("{0}_", _protocolName);
            }
        }

        public static string GetCommandUrlStatic(string link, string pageId, string objectId)
        {
            return string.Concat(
                        link.Replace(Constants.OneNoteProtocol, ProtocolFullString),
                        "&", Constants.QueryParameterKey_CustomPageId, "=", pageId,
                        "&", Constants.QueryParameterKey_CustomObjectId, "=", objectId);
        }

        public string ProtocolName
        {
            get { return _protocolName; }
        }

        public bool IsProtocolCommand(string[] args)
        {
            return args.Length > 0 && args[0].StartsWith(ProtocolName, StringComparison.OrdinalIgnoreCase);
        }

        public void ExecuteCommand(string[] args)
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

        private void TryExecuteCommand(string[] args)
        {
            if (args.Length == 0)
                throw new ArgumentNullException("args");

            var oneNoteApp = new Microsoft.Office.Interop.OneNote.Application();

            try
            {
                var newPath = args[0].Replace(ProtocolFullString, Constants.OneNoteProtocol);

                if (!TryToRedirectByIds(oneNoteApp, newPath))
                {
                    if (!TryToRedirectByUrl(oneNoteApp, newPath))
                        throw new Exception(string.Format("The {0} attempts of NavigateToUrl() were unsuccessful.", NavigateAttemptsCount));         
                }
            }
            finally
            {
                Marshal.ReleaseComObject(oneNoteApp);
                oneNoteApp = null;
            }
        }

        private static bool TryToRedirectByUrl(Microsoft.Office.Interop.OneNote.Application oneNoteApp, string newPath)
        {
            var currentPageId = oneNoteApp.Windows.CurrentWindow.CurrentPageId;
            newPath = GetValidPath(newPath);
            for (int i = 0; i < NavigateAttemptsCount; i++)
            {
                try
                {
                    if (currentPageId == oneNoteApp.Windows.CurrentWindow.CurrentPageId)
                        oneNoteApp.NavigateToUrl(newPath);

                    return true;
                }
                catch (COMException ex)
                {
                    //if (ex.Message.IndexOf("0x80042014") != -1)  //hrObjectDoesNotExist
                    //    return true;

                    Thread.Sleep(1000);
                }
            }

            return false;
        }

        private static bool TryToRedirectByIds(Microsoft.Office.Interop.OneNote.Application oneNoteApp, string newPath)
        {
            var pageId = StringUtils.GetNotFramedAttributeValue(newPath, Constants.QueryParameterKey_CustomPageId);

            if (!string.IsNullOrEmpty(pageId))
            {
                var objectId = !string.IsNullOrEmpty(pageId) 
                                    ? StringUtils.GetNotFramedAttributeValue(newPath, Constants.QueryParameterKey_CustomObjectId) 
                                    : string.Empty;
                try
                {
                    oneNoteApp.NavigateTo(pageId, objectId);
                    return true;
                }
                catch (COMException)
                {
                }
            }

            return false;
        }

        private static string GetValidPath(string newPath)
        {
            return string.Format("{0}//{1}{2}",
                Constants.OneNoteProtocol,
                Path.GetPathRoot(Environment.SystemDirectory),
                newPath.Substring(Constants.OneNoteProtocol.Length + 5));
        }

        string IProtocolHandler.GetCommandUrl(string args)
        {
            return string.Format("{0}:{1}", ProtocolName, args);
        }  
    }
}
