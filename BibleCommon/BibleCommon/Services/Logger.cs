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
        public static bool ErrorWasLogged = false;        
        private static int _level = 0;
        private static FileStream _fileStream = null;
        private static StreamWriter _streamWriter = null;
        private static ListBox _lb = null;

        public static List<string> Errors { get; set; }

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
                string directoryPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), Consts.Constants.ToolsName);

                if (!Directory.Exists(directoryPath))
                    Directory.CreateDirectory(directoryPath);

                _fileStream = new FileStream(Path.Combine(directoryPath, systemName + ".txt"), FileMode.Create);
                _streamWriter = new StreamWriter(_fileStream);

                Errors = new List<string>();

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
                if (_streamWriter != null)
                    _streamWriter.Close();

                if (_fileStream != null)
                    _fileStream.Close();

                _isInitialized = false;
                //_lb = null;
            }
        }

        public static void LogMessage(string message, bool leveled, bool newLine, bool writeDateTime = true)
        {
            LogMessageToFileAndConsole(false, string.Empty, null, writeDateTime, false);

            if (leveled)
                for (int i = 0; i < _level; i++)
                    LogMessageToFileAndConsole(false, "  ", null, false, false);

            LogMessageToFileAndConsole(newLine, message, null, false, false);
        }


        private static bool _newLineForListBox = false;
        private static void LogMessageToFileAndConsole(bool newLine, string message, string messageEx, bool writeDateTime, bool isError)
        {
            if (string.IsNullOrEmpty(messageEx))
                messageEx = message;

            if (writeDateTime)
                messageEx = string.Format("{0}: {1}", DateTime.Now, messageEx);


            if (_lb != null)
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
                

                if (_streamWriter != null && _streamWriter.BaseStream != null)
                    _streamWriter.WriteLine(messageEx);                    
            }
            else
            {
                Console.Write(message);
                if (_lb != null)
                    _newLineForListBox = false;

                if (_streamWriter != null && _streamWriter.BaseStream != null)
                    _streamWriter.Write(messageEx);
            }

            if (_streamWriter != null && _streamWriter.BaseStream != null)
                _streamWriter.Flush();

            if (isError)
                if (Errors != null)
                    Errors.Add(message);
        }     

        public static void LogMessage(string message, params object[] args)
        {
            LogMessage(FormatString(message, args), true, true);
        }

        public static void LogError(string message, Exception ex)
        {
            LogMessageToFileAndConsole(true, string.Format("{0}{1} {2}", _errorText, message, ex.Message), string.Format("{0} {1}", message, ex.ToString()), true, true);
            ErrorWasLogged = true;   
        }

        public static void LogError(Exception ex)
        {
            LogError(string.Empty, ex);
        }

        public static void LogError(string message, params object[] args)
        {
            LogMessageToFileAndConsole(true, _errorText + FormatString(message, args), null, true, true);
            ErrorWasLogged = true;
        }

        private static string FormatString(string message, params object[] args)
        {
            return args.Count() == 0 ? message : string.Format(message, args);
        }
    }
}
