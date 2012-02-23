using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BibleCommon.Common
{
    public class NotFoundVerseLinkPageExceptions : Exception
    {
        public NotFoundVerseLinkPageExceptions(string message)
            : base(message)
        {
        }
    }
}
