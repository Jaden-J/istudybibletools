using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BibleVerseLinker
{
    public static class Logger
    {
        public static bool WasLogged = false;

        public static void LogError(string message)
        {
            Console.WriteLine(message);
            WasLogged = true;
        }
    }
}
