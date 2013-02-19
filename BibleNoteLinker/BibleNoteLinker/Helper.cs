using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BibleCommon.Services;
using BibleNoteLinker.Properties;

namespace BibleNoteLinker
{
    public static class Helper
    {
        public static string GetRightFoundPagesString(int pagesCount)
        {
            string firstPart = pagesCount == 1 ? BibleCommon.Resources.Constants.NoteLinkerOneFound : BibleCommon.Resources.Constants.NoteLinkerManyFound;

            return string.Format("{0}: {1}", firstPart, GetRightPagesString(pagesCount));
        }


        public static string GetRightPagesString(int pagesCount)
        {
            string s = BibleCommon.Resources.Constants.NoteLinkerOfManyPages;
            int tempPagesCount = pagesCount;

            tempPagesCount = tempPagesCount % 100;
            if (!(tempPagesCount >= 10 && tempPagesCount <= 20))
            {
                tempPagesCount = tempPagesCount % 10;

                if (tempPagesCount == 1)
                    s = BibleCommon.Resources.Constants.NoteLinkerOnePage;
                else if (tempPagesCount >= 2 && tempPagesCount <= 4)
                    s = BibleCommon.Resources.Constants.NoteLinkerManyPages;
            }            

            return string.Format("{0} {1}", pagesCount, s);
        }

    }
}
