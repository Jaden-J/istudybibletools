using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BibleCommon.Resources;

namespace BibleCommon.Common
{
    public class NotConfiguredException : Exception
    {
        public NotConfiguredException()
            : base(BibleCommon.Resources.Constants.Error_SystemIsNotConfigures)
        {
        }

        public NotConfiguredException(string message)
            : base(string.Format("{0} {1}", BibleCommon.Resources.Constants.Error_SystemIsNotConfigures, message))
        {
        }
    }

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
        public enum Severity
        {
            Warning,
            Error
        }

        public Severity Level { get; set; }

        public BaseVersePointerException(string message, Severity level)
            : base(message)
        {
            this.Level = level;
        }
    }

    public class ParallelVerseNotFoundException : BaseVersePointerException
    {
        public ParallelVerseNotFoundException(SimpleVersePointer verse, Severity level)
            : base(string.Format("There is no parallel verse '{0}'", verse), level)
        {
        }
    }

    public class GetParallelVerseException : BaseVersePointerException
    {
        public GetParallelVerseException(string message, SimpleVersePointer baseVerse, Severity level)
            : base(string.Format("Can not find parallel verse for baseVerse '{0}': {1}", baseVerse.ToString(), message), level)
        {
        }
    }
}
