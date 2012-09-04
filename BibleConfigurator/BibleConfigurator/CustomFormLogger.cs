using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BibleCommon.Contracts;

namespace BibleConfigurator
{
    public class CustomFormLogger: ICustomLogger
    {
        public void LogMessage(string message, params object[] args)
        {
            throw new NotImplementedException();
        }

        public void LogWarning(string message, params object[] args)
        {
            throw new NotImplementedException();
        }

        public void LogException(string message, params object[] args)
        {
            throw new NotImplementedException();
        }
    }
}
