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
using Microsoft.Office.Interop.OneNote;
using System.Text.RegularExpressions;

namespace BibleCommon.Handlers
{
    public class NavigateToHandler : IProtocolHandler
    {
        private const string _protocolName = "isbtOpen:";
        private const int NavigateAttemptsCount = 3;

        public static string ProtocolFullString
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

        public bool IsProtocolCommand(params string[] args)
        {
            return args.Length > 0 && args[0].StartsWith(ProtocolName, StringComparison.OrdinalIgnoreCase);
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

        private void TryExecuteCommand(string[] args)
        {
            if (args.Length == 0)
                throw new ArgumentNullException("args");

            var oneNoteApp = new Microsoft.Office.Interop.OneNote.Application();

            try
            {
                var newPath = args[0].ReplaceIgnoreCase(ProtocolFullString, Constants.OneNoteProtocol);                

                if (!TryToRedirectByIds(ref oneNoteApp, newPath))
                {
                    if (!TryToRedirectByUrl(ref oneNoteApp, newPath))
                    {
                        //throw new Exception(string.Format("OneNote cannot open the specified location after {0} attempts: {1}", NavigateAttemptsCount, newPath));
                    }
                }
            }
            finally
            {
                Marshal.ReleaseComObject(oneNoteApp);
                oneNoteApp = null;
            }
        }

        private static bool TryToRedirectByUrl(ref Application oneNoteApp, string newPath)
        {            
            var pageId = oneNoteApp.Windows.CurrentWindow != null ? oneNoteApp.Windows.CurrentWindow.CurrentPageId : null;
            newPath = GetValidPath(newPath);
            for (int i = 0; i < NavigateAttemptsCount; i++)
            {
                try
                {
                    var currentPageId = oneNoteApp.Windows.CurrentWindow != null ? oneNoteApp.Windows.CurrentWindow.CurrentPageId : null;
                    if (pageId == currentPageId)
                    {
                        OneNoteUtils.UseOneNoteAPI(ref oneNoteApp, (oneNoteAppSafe) =>
                        {
                            oneNoteAppSafe.NavigateToUrl(newPath);
                        });
                    }

                    return true;
                }
                catch (COMException ex)
                {
                    //if (ex.Message.IndexOf("0x80042014") != -1)  //hrObjectDoesNotExist
                    //    return true;

                    Thread.Sleep(2000);
                }
            }

            return false;
        }

        private static bool TryToRedirectByIds(ref Application oneNoteApp, string newPath)
        {            
            var pageId = StringUtils.GetNotFramedAttributeValue(newPath, Constants.QueryParameterKey_CustomPageId);

            if (!string.IsNullOrEmpty(pageId))
            {
                var objectId = !string.IsNullOrEmpty(pageId) 
                                    ? StringUtils.GetNotFramedAttributeValue(newPath, Constants.QueryParameterKey_CustomObjectId) 
                                    : string.Empty;
                try
                {
                    OneNoteUtils.UseOneNoteAPI(ref oneNoteApp, (oneNoteAppSafe) =>
                    {
                        oneNoteAppSafe.NavigateTo(pageId, objectId);
                    });
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
            return Regex.Replace(newPath, @"//\D:\\", "//" + Path.GetPathRoot(Environment.SystemDirectory));            
        }

        string IProtocolHandler.GetCommandUrl(string args)
        {
            return string.Format("{0}:{1}", ProtocolName, args);
        }  
    }
}
