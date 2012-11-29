using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BibleCommon.Contracts
{
    public interface IProtocolHandler
    {
        string ProtocolName { get; }
        string GetCommandUrl(string args);
        bool IsProtocolCommand(string[] args);
        bool ExecuteCommand(string[] args);
    }
}
