using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using BibleCommon.Helpers;
using System.Windows.Forms;
using System.Drawing;

namespace BibleCommon.Services
{
    public static class Logger
    {
        public enum Severity
        {
            Info,
            Warning,
            Error            
        }

        public static bool ErrorWasLogged = false;
        public static bool WarningWasLogged = false;        
        private static int _level = 0;
        private static string _logFilePath;
        private static FileStream _fileStream = null;
        private static StreamWriter _streamWriter = null;
        private static ListBox _lb = null;

        public static List<string> Errors { get; set; }
        public static List<string> Warnings { get; set; }

        public static string LogFilePath
        {
            get
            {
                return _logFilePath;
            }
        }

        private static string _errorText = BibleCommon.Resources.Constants.ErrorUpper + ": ";

        public static void MoveLevel(int levelDiv)
        {
            _level += levelDiv;
        }

        private static bool _isInitialized = false;

        public static void Init(string systemName)
        {
            if (!_isInitialized)
            {
                string directoryPath = Path.Combine(
                                            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), Consts.Constants.ToolsName), 
                                            Consts.Constants.LogsDirectory);

                if (!Directory.Exists(directoryPath))
                    Directory.CreateDirectory(directoryPath);

                _logFilePath = Path.Combine(directoryPath, systemName + ".txt");
                try
                {
                    _fileStream = new FileStream(_logFilePath, FileMode.Create);
                }
                catch (IOException)
                {
                    _logFilePath = Path.Combine(directoryPath, string.Format("{0}_{1}.txt", systemName, Guid.NewGuid()));
                    _fileStream = new FileStream(_logFilePath, FileMode.Create);
                }

                _streamWriter = new StreamWriter(_fileStream, Encoding.UTF8);

                Errors = new List<string>();
                Warnings = new List<string>();

                _isInitialized = true;
            }
        }

        public static void SetOutputListBox(ListBox lb)
        {
            _lb = lb;
        }

        public static void Done()
        {
            if (_isInitialized)
            {
                bool needToDelete = false;

                if (_fileStream != null)
                {
                    if (_fileStream.Length == 0)
                        needToDelete = true;
                    _fileStream.Close();
                }                            


                _isInitialized = false;

                if (needToDelete)
                {
                    try
                    {
                        File.Delete(_logFilePath);
                    }
                    catch { }
                }
                //_lb = null;
            }
        }

        public static void LogMessage(string message, bool leveled, bool newLine, bool writeDateTime = true, bool silient = false, Severity severity = Severity.Info)
        {
            LogMessageToFileAndConsole(false, string.Empty, null, writeDateTime, silient, severity);

            if (leveled)
                for (int i = 0; i < _level; i++)
                    LogMessageToFileAndConsole(false, "  ", null, false, silient, severity);

            LogMessageToFileAndConsole(newLine, message, null, false, silient, severity);
        }


        private static bool _newLineForListBox = false;
        private static void LogMessageToFileAndConsole(bool newLine, string message, string messageEx, bool writeDateTime, bool silient, Severity severity)
        {
            if (!_isInitialized)
            {
                try
                {
                    Init(System.Reflection.Assembly.GetEntryAssembly().GetName().Name);
                }
                catch { }
            }            

            if (string.IsNullOrEmpty(messageEx))
                messageEx = message;

            if (writeDateTime)
                messageEx = string.Format("{0}: {1}", DateTime.Now, messageEx);

            if (_lb != null && !silient)
            {
                if (_newLineForListBox || _lb.Items.Count == 0)
                {
                    _lb.Items.Add(message);
                    _lb.SelectedIndex = _lb.Items.Count - 1;
                }
                else
                    _lb.Items[_lb.Items.Count - 1] += message;


                int width = Convert.ToInt32(message.Length * 5.75);
                if (width > _lb.HorizontalExtent)
                    _lb.HorizontalExtent = width;
            }

            if (newLine)
            {
                Console.WriteLine(message);
                if (_lb != null)
                    _newLineForListBox = true;

                TryToWriteToFile(messageEx, newLine);
            }
            else
            {
                Console.Write(message);
                if (_lb != null)
                    _newLineForListBox = false;

                TryToWriteToFile(messageEx, newLine);
            }

            try
            {
                if (_streamWriter != null && _streamWriter.BaseStream != null)
                    _streamWriter.Flush();
            }
            catch { }

            if (severity == Severity.Error)
            {
                if (Errors != null)
                    Errors.Add(message);
            }
            else if (severity == Severity.Warning)
            {
                if (Warnings != null)
                    if (!string.IsNullOrEmpty(message) && !string.IsNullOrEmpty(message.Trim()))
                        Warnings.Add(message);
            }
        }

        private static void TryToWriteToFile(string message, bool newLine)
        {
            try
            {
                if (_streamWriter != null && _streamWriter.BaseStream != null)
                {
                    if (newLine)
                        _streamWriter.WriteLine(message);
                    else
                        _streamWriter.Write(message);
                }
            }
            catch (Exception subEx)
            {
                MessageBox.Show(subEx.ToString());
                MessageBox.Show(message);
            }
        }


        public static void LogWarning(string message, params object[] args)
        {
            LogMessage("Warning: " + FormatString(message, args), true, true, true, false, Severity.Warning);
            WarningWasLogged = true;
        }

        /// <summary>
        /// Log only to log
        /// </summary>
        /// <param name="message"></param>
        /// <param name="args"></param>
        public static void LogMessageSilientParams(string message, params object[] args)
        {
            LogMessage(FormatString(message, args), true, true, true, true);
        }

        public static void LogMessageParams(string message, params object[] args)
        {
            LogMessage(FormatString(message, args), true, true);
        }

        public static void LogError(string message, Exception ex)
        {
            LogMessageToFileAndConsole(true, string.Format("{0}{1} {2}", _errorText, message, OneNoteUtils.ParseError(ex.Message)), string.Format("{0} {1}", message, ex.ToString()), true, false, Severity.Error);
            ErrorWasLogged = true;   
        }

        public static void LogError(Exception ex)
        {
            LogError(string.Empty, ex);
        }

        public static void LogError(string message, params object[] args)
        {
            LogMessageToFileAndConsole(true, _errorText + FormatString(message, args), null, true, false, Severity.Error);
            ErrorWasLogged = true;
        }

        private static string FormatString(string message, params object[] args)
        {
            return args.Count() == 0 ? message : string.Format(message, args);
        }
    }
}
