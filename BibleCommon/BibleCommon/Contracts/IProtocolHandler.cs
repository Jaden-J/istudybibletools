using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BibleCommon.Contracts
{
    public interface IProtocolHandler
    {
        string GetCommandUrl(string args);
        void ExecuteCommand(string args);
    }
}
