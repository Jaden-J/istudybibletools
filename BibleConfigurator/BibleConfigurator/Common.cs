using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BibleConfigurator
{
    public enum NotebookType
    {
        Single,
        Bible,
        BibleComments,
        BibleStudy
    }

    public enum SectionGroupType
    {
        Bible,
        BibleComments,
        BibleStudy
    }

    public class SectionGroupInfo
    {
        public SectionGroupType Type { get; set; }
        public string Id { get; set; }
        public string OriginalName { get; set; }
        public string NewName { get; set; }
    }

    public class InvalidNotebookException: Exception
    {
    }

    public class LoadParametersException : Exception
    {
    }
}
