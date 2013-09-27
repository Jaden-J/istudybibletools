using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BibleCommon.Consts
{
    public static class SystemConstants
    {
        public static readonly bool IsOneNote2010 = true;

        public enum OneNoteVersion
        {
            v2010 = 2010,
            v2013 = 2013
    }

        public static OneNoteVersion VersionOfOneNote
        {
            get
            {
                if (IsOneNote2010)
                    return OneNoteVersion.v2010;
                else
                    return OneNoteVersion.v2013;
            }
        }
    }
}
