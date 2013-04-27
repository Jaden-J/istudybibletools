using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BibleCommon.Resources;

namespace BibleCommon.Common
{
    public class ProgramException : Exception
    {
        public ProgramException(string message, params string[] args)
            : base(string.Format(message, args))
        {
        }
    }

    public class NotConfiguredException : ProgramException
    {
        public NotConfiguredException()
            : base(BibleCommon.Resources.Constants.Error_SystemIsNotConfigured)
        {
        }

        public NotConfiguredException(string message)
            : base(BibleCommon.Resources.Constants.Error_SystemIsNotConfigured, message)
        {
        }
    }

    public class NotFoundPageException : ProgramException
    {
        public NotFoundPageException(string message)
            : base(message)
        {
        }
    }

    public class NotFoundVerseLinkPageException : NotFoundPageException
    {
        public NotFoundVerseLinkPageException(string message)
            : base(message)
        {
        }        
    }

    public class InvalidModuleException : ProgramException
    {
        public InvalidModuleException(string message)
            : base(Constants.Error_InvalidModule + " " + message)
        {
        }
    }

    public class ModuleNotFoundException : InvalidModuleException
    {
        public ModuleNotFoundException(string message)
            : base(Constants.Error_InvalidModule + " " + message)
        {
        }
    }

    public class InvalidNotebookException : ProgramException
    {
        public InvalidNotebookException(string message)
            : base(message)
        {
        }
    }

    public class BaseVersePointerException : ProgramException
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

    public class ChapterNotFoundException : BaseVersePointerException
    {
        public ChapterNotFoundException(SimpleVersePointer verse, string moduleShortName, Severity level)
            : base(string.Format("There is no chapter '({2}) {0} {1}'", verse.BookIndex, verse.Chapter, moduleShortName), level)
        {
            this.IsChapterException = true;
        }
    }

   
}
