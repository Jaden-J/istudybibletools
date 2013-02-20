using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using ISBTCommandHandler.SingleInstanceService;
using System.Reflection;
using BibleCommon.Contracts;
using BibleCommon.Handlers;

namespace ISBTCommandHandler
{
    static class Program
    {
        private static CommandForm _mainForm;

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main(params string[] args)
        {
            try
            {
                if (!ProcessCommandWithSimpleHandler(args))
                {
                    if (!ApplicationInstanceManager.CreateSingleInstance(
                                                        Assembly.GetExecutingAssembly().GetName().Name,
                                                        SingleInstanceCallback))
                        return;

                    _mainForm = new CommandForm();
                    Application.Run(_mainForm);
                }
            }
            catch (Exception ex) 
            {
                LogError(ex, args);
            }           
        }

        private static bool ProcessCommandWithSimpleHandler(string[] args)
        {            
            var simpleHandlers = new IProtocolHandler[] { new NavigateToHandler() };   // то есть хэндлеры, для которых не нужен кэш

            foreach (var simpleHandler in simpleHandlers)
            {
                if (simpleHandler.IsProtocolCommand(args))
                {
                    simpleHandler.ExecuteCommand(args);
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Single instance callback handler.
        /// </summary>
        /// <param name="sender">The sender.</param>
        /// <param name="args">The <see cref="SingleInstanceApplication.InstanceCallbackEventArgs"/> instance containing the event data.</param>
        private static void SingleInstanceCallback(object sender, InstanceCallbackEventArgs args)
        {
            if (args == null || _mainForm == null) return;
            Action<bool> d = (bool x) =>
            {
                _mainForm.ProcessCommandLine(args.CommandLineArgs);                
            };
            _mainForm.Invoke(d, true);
        }

        internal static void LogError(Exception ex, params string[] args)
        {
            var directoryPath = Path.Combine(
                                            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "IStudyBibleTools"),
                                            "Logs");

            if (!Directory.Exists(directoryPath))
                Directory.CreateDirectory(directoryPath);

            var logFilePath = Path.Combine(directoryPath, "ISBTCommandHandler.txt");

            File.AppendAllText(logFilePath, string.Format("args: {0}, \nException: {1}\n", string.Join(";\t", args), ex.ToString()));
        }
    }
}
