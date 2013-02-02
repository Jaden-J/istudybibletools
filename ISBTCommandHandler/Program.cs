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

                    for (int i = 0; i < 10; i++)
                    {
                        try
                        {
                            oneNoteApp.NavigateToUrl(newPath);
                            break;
                        }
                        catch (COMException)
                        {
                            Thread.Sleep(1000);
                        }                        
                    }

                    var currentPageId = oneNoteApp.Windows.CurrentWindow.CurrentPageId;
                    Thread.Sleep(500);

                    for (int i = 0; i < 10; i++)
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

                    throw new Exception("The 10 attempts of NavigateToUrl() were unsuccessful.");

                    //var parts = Uri.UnescapeDataString(args[0].Split(new char[] { ':' })[1]).Split(new char[] { ';' });

                    //oneNoteApp.NavigateTo(parts[0], parts[1]);                    
                }                
                catch (Exception ex)
                {

                    string directoryPath = Path.Combine(
                                                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "IStudyBibleTools"),
                                                "Logs");

                    if (!Directory.Exists(directoryPath))
                        Directory.CreateDirectory(directoryPath);

                    var logFilePath = Path.Combine(directoryPath, "ISBTCommandHandler.txt");

                    File.AppendAllText(logFilePath, string.Format("args: {0}, \nException: {1}\n", string.Join(";\t", args), ex.ToString()));
                }
            }
        }
    }
}
