using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BibleCommon.Contracts
{
    public interface ICustomLogger
    {
        bool AbortedByUser { get; }

        void LogMessage(string message, params object[] args);
        void LogWarning(string message, params object[] args);
        void LogException(string message, params object[] args);
    }
}
