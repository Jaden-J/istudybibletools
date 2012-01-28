using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using BibleCommon.Helpers;

namespace BibleCommon.Services
{
    public static class Logger
    {
        public static bool ErrorWasLogged = false;        
        private static int _level = 0;
        private static FileStream _fileStream = null;
        private static StreamWriter _streamWriter = null;



        public static void MoveLevel(int levelDiv)
        {
            _level += levelDiv;
        }

        public static void Init(string systemName)
        {

            string directoryPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), Consts.Constants.ToolsName);

            if (!Directory.Exists(directoryPath))
                Directory.CreateDirectory(directoryPath);

            _fileStream = new FileStream(Path.Combine(directoryPath, systemName + ".txt"), FileMode.Create);
            _streamWriter = new StreamWriter(_fileStream);
        }

        public static void Done()
        {
            if (_streamWriter != null)
                _streamWriter.Close();

            if (_fileStream != null)
                _fileStream.Close();
        }

        public static void LogMessage(string message, bool leveled, bool newLine, bool writeDateTime = true)
        {
            LogMessageToFileAndConsole(false, string.Empty, null, writeDateTime);

            if (leveled)
                for (int i = 0; i < _level; i++)
                    LogMessageToFileAndConsole(false, "  ", null, false);

            LogMessageToFileAndConsole(newLine, message, null, false);
        }

        private static void LogMessageToFileAndConsole(bool newLine, string message, string messageEx = null, bool writeDateTime = true)
        {
            if (string.IsNullOrEmpty(messageEx))
                messageEx = message;

            if (writeDateTime)
                messageEx = string.Format("{0}: {1}", DateTime.Now, messageEx);

            if (newLine)
            {
                Console.WriteLine(message);                

                if (_streamWriter != null && _streamWriter.BaseStream != null)
                    _streamWriter.WriteLine(messageEx);                    
            }
            else
            {
                Console.Write(message);
                if (_streamWriter != null && _streamWriter.BaseStream != null)
                    _streamWriter.Write(messageEx);
            }

            if (_streamWriter != null && _streamWriter.BaseStream != null)
                _streamWriter.Flush();
        }

        public static void LogMessage(string message, params object[] args)
        {
            LogMessage(string.Format(message, args), true, true);
        }

        public static void LogError(string message, Exception ex)
        {
            LogMessageToFileAndConsole(true, string.Format("ОШИБКА: {0} {1}", message, ex.Message), string.Format("{0} {1}", message, ex.ToString()));
            ErrorWasLogged = true;   
        }

        public static void LogError(Exception ex)
        {
            LogError(string.Empty, ex);
        }

        public static void LogError(string message, params object[] args)
        {            
            LogMessageToFileAndConsole(true, "ОШИБКА: " + string.Format(message, args));
            ErrorWasLogged = true;
        }
    }
}
