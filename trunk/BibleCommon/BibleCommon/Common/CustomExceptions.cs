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

    public abstract class BaseVersePointerException: Exception
    {
        public BaseVersePointerException(string message)
            : base(message)
        { 
        }
    }

    public class VerseNotFoundException : BaseVersePointerException
    {
        public VerseNotFoundException(SimpleVersePointer verse)
            : base(string.Format("There is no verse '{0}'", verse))
        {
        }
    }

    public class GetParallelVerseException : BaseVersePointerException
    {
        public GetParallelVerseException(string message, SimpleVersePointer baseVerse)
            : base(string.Format("Can not find parallel verse for baseVerse '{0}': {1}", baseVerse.ToString(), message))
        {
        }
    }
}
