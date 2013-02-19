using Microsoft.Office.Interop.OneNote;
using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;

namespace ISBTCommandHandler
{
    static class Program
    {
        const string QueryParameterKey_CustomPageId = "cpId";
        const string QueryParameterKey_CustomObjectId = "coId";
        const int NavigateAttemptsCount = 10;
        const string OneNoteProtocol = "onenote:";

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main(params string[] args)
        {            
            if (args.Length == 1)
            {
                string newPath = null;
                Application oneNoteApp = null;
                try
                {
                    oneNoteApp = new Application();
                    newPath = args[0].Replace("isbtopen:_", OneNoteProtocol);

                    if (!TryToRedirectByIds(oneNoteApp, newPath))
                    {
                        if (!TryToRedirectByUrl(oneNoteApp, newPath))
                            throw new Exception(string.Format("The {0} attempts of NavigateToUrl() were unsuccessful.", NavigateAttemptsCount));                        
                    }
                }                
                catch (Exception ex)
                {
                    LogError(ex, args);                    
                }
            }            
        }

        private static bool TryToRedirectByUrl(Application oneNoteApp, string newPath)
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
                    if (ex.Message.IndexOf("0x80042014") != -1)  //hrObjectDoesNotExist
                        return true;

                    Thread.Sleep(1000);
                }
            }

            return false;
        }

        private static bool TryToRedirectByIds(Application oneNoteApp, string newPath)
        {
            var pageId = GetAttributeValue(newPath, QueryParameterKey_CustomPageId);

            if (!string.IsNullOrEmpty(pageId))
            {
                var objectId = !string.IsNullOrEmpty(pageId) ? GetAttributeValue(newPath, QueryParameterKey_CustomObjectId) : string.Empty;
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
                OneNoteProtocol, 
                Path.GetPathRoot(Environment.SystemDirectory),
                newPath.Substring(OneNoteProtocol.Length + 5));                
        }

        private static void LogError(Exception ex, params string[] args)
        {
            var directoryPath = Path.Combine(
                                            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "IStudyBibleTools"),
                                            "Logs");

            if (!Directory.Exists(directoryPath))
                Directory.CreateDirectory(directoryPath);

            var logFilePath = Path.Combine(directoryPath, "ISBTCommandHandler.txt");

            File.AppendAllText(logFilePath, string.Format("args: {0}, \nException: {1}\n", string.Join(";\t", args), ex.ToString()));
        }

        private static string GetAttributeValue(string s, string attributeName)
        {
            string result = null;

            var searchString = string.Format("&{0}=", attributeName);

            var startIndex = s.IndexOf(searchString);

            if (startIndex > -1)
            {
                startIndex = startIndex + searchString.Length;

                if (s.Length > startIndex)
                {
                    int endIndex = s.IndexOf("&", startIndex + 1);

                    if (endIndex > -1)
                        result = s.Substring(startIndex, endIndex - startIndex);
                    else
                        result = s.Substring(startIndex);
                }
            }
            return result;
        }           
    }
}
