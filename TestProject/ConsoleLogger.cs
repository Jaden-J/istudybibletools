using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BibleCommon.Contracts;

namespace TestProject
{
    public class ConsoleLogger: ICustomLogger
    {
        public void LogMessage(string message, params object[] args)
        {
            Console.WriteLine(message, args);
        }

        public void LogWarning(string message, params object[] args)
        {
            Console.WriteLine(message, args);
        }

        public void LogException(string message, params object[] args)
        {
            Console.WriteLine(message, args);
        }

        public bool AbortedByUser
        {
            get { return false; }
        }
    }
}
