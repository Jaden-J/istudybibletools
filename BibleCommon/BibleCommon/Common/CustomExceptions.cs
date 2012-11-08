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

    public class InvalidNotebookException : Exception
    {
        public InvalidNotebookException(string message)
            : base(message)
        {
        }
    }

    public class BaseVersePointerException: Exception
    {
        public enum Severity
        {
            Warning,
            Error
        }

        public Severity Level { get; set; }

        public bool IsChapterException { get; set; }

        public BaseVersePointerException(string message, Severity level)
            : base(message)
        {
            this.Level = level;
        }
    }

    public class VerseNotFoundException : BaseVersePointerException
    {
        public VerseNotFoundException(SimpleVersePointer verse, string moduleShortName, Severity level)
            : base(string.Format("There is no verse '({1}) {0}'", verse, moduleShortName), level)
        {
        }
    }

    public class GetParallelVerseException : BaseVersePointerException
    {
        public GetParallelVerseException(string message, SimpleVersePointer baseVerse, string moduleShortName, Severity level)
            : base(string.Format("Can not find parallel verse for baseVerse '({1}) {0}': {2}", baseVerse, moduleShortName, message), level)
        {
        }
    }

    public class BaseChapterSectionNotFoundException : BaseVersePointerException
    {
        public BaseChapterSectionNotFoundException(int baseChapterIndex, int baseBookInfoIndex)
            : base(string.Format("Can not find the page for chapter '{0} {1}' in base Bible", baseBookInfoIndex, baseChapterIndex), Severity.Error)
        {
            this.IsChapterException = true;
        }
    }

    public class ChapterNotFoundException: BaseVersePointerException
    {
        public ChapterNotFoundException(SimpleVersePointer verse, string moduleShortName, Severity level)
            : base(string.Format("There is no chapter '({2}) {0} {1}'", verse.BookIndex, verse.Chapter, moduleShortName), level)
        {
            this.IsChapterException = true;
        }
    }

   
}
