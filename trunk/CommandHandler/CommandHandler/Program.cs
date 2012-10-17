using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace CommandHandler
{
    static class Program
    {
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            var openVerseHandler = new NavigateToOneNoteHandler();

            if (openVerseHandler.IsProtocolCommand(args))
                openVerseHandler.ExecuteCommand(args);
        }
    }
}
