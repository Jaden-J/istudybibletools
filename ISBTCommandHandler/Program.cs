using Microsoft.Office.Interop.OneNote;
using System;
using System.IO;

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
                try
                {
                    Application oneNoteApp = new Application();

                    var parts = Uri.UnescapeDataString(args[0].Split(new char[] { ':' })[1]).Split(new char[] { ';' });

                    oneNoteApp.NavigateTo(parts[0], parts[1]);
                }
                catch (Exception ex)
                {
                    string directoryPath = Path.Combine(
                                                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "IStudyBibleTools"),
                                                "Logs");

                    if (!Directory.Exists(directoryPath))
                        Directory.CreateDirectory(directoryPath);

                    var logFilePath = Path.Combine(directoryPath, "ISBTCommandHandler.txt");

                    File.AppendAllText(logFilePath, string.Format("args: {0}, \nException: {1}", string.Join(";\t", args),  ex.ToString()));
                }
            }
        }
    }
}
