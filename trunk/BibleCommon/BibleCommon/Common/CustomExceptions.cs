using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BibleCommon.Resources;

namespace BibleCommon.Common
{
    public class NotFoundVerseLinkPageExceptions : Exception
    {
        public NotFoundVerseLinkPageExceptions(string message)
            : base(message)
        {
        }        
    }

    public class InvalidModuleException : Exception
    {
        public InvalidModuleException(string message)
            : base(Constants.Error_InvalidModule + " " + message)
        {
        }     
    }
}
