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
                    newPath = args[0].Replace("isbtopen:_", "onenote:");
                    var pageId = GetAttributeValue(newPath, QueryParameterKey_CustomPageId);
                    var objectId = !string.IsNullOrEmpty(pageId) ? GetAttributeValue(newPath, QueryParameterKey_CustomObjectId) : string.Empty;

                    if (!string.IsNullOrEmpty(pageId))
                    {
                        try
                        {
                            oneNoteApp.NavigateTo(pageId, objectId);
                            return;
                        }
                        catch (COMException)
                        {
                        }
                    }

                    //если дошли до сюда, значит не удалось перейти выше
                    var currentPageId = oneNoteApp.Windows.CurrentWindow.CurrentPageId;
                    for (int i = 0; i < NavigateAttemptsCount; i++)
                    {
                        try
                        {
                            if (currentPageId == oneNoteApp.Windows.CurrentWindow.CurrentPageId)                                
                                oneNoteApp.NavigateToUrl(newPath);

                            break;
                        }
                        catch (COMException)
                        {
                            Thread.Sleep(1000);
                        }                        
                    }

                    throw new Exception(string.Format("The {0} attempts of NavigateToUrl() were unsuccessful.", NavigateAttemptsCount));
                }                
                catch (Exception ex)
                {
                    LogError(ex, args);                    
                }
            }
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

            string searchString = string.Format("{0}=", attributeName);

            int startIndex = s.IndexOf(searchString);

            startIndex = startIndex + searchString.Length;

            if (s.Length > startIndex)
            {
                int endIndex = s.IndexOf("&", startIndex + 1);

                if (endIndex > -1)                
                    result = s.Substring(startIndex, endIndex - startIndex);
                else
                    result = s.Substring(startIndex);                
            }
            return result;
        }           
    }
}
