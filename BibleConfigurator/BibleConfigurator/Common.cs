using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BibleCommon.Common;

namespace BibleConfigurator
{
    public class SectionGroupDTO
    {
        public ContainerType Type { get; set; }
        public string Id { get; set; }
        public string OriginalName { get; set; }
        public string NewName { get; set; }
    }

    public class InvalidNotebookException: Exception
    {
        public InvalidNotebookException(string message, params object[] args)
            : base(string.Format(message, args))
        {
        }

        public InvalidNotebookException()
        {
        }
    }

    public class SaveParametersException : Exception
    {
        public bool NeedToReload { get; set; }

        public SaveParametersException(string message, bool needToReload)
            : base(message)
        {
            this.NeedToReload = needToReload;
        }
    }    
}
