using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BibleConfigurator
{
    public static class Consts
    {
        public const string SingleNotebookDefaultName = "Holy Bible";
        public const string BibleNotebookDefaultName = "Библия";
        public const string BibleCommentsNotebookDefaultName = "Комментарии к Библии";
        public const string BibleStudyNotebookDefaultName = "Изучение Библии";

        public const string SingleNotebookTemplateFileName = "Holy Bible.onepkg";
        public const string BibleNotebookTemplateFileName = "Библия.onepkg";
        public const string BibleCommentsNotebookTemplateFileName = "Комментарии к Библии.onepkg";
        public const string BibleStudyNotebookTemplateFileName = "Изучение Библии.onepkg";
        public const string TemplatesDirectory = "OneNoteTemplates";

        public const string OldTestamentName = "Ветхий Завет";
        public const string NewTestamentName = "Новый Завет";

        public const string BibleSectionGroupDefaultName = "Библия";
        public const string BibleCommentsSectionGroupDefaultName = "Комментарии к Библии";
        public const string BibleStudySectionGroupDefaultName = "Изучение Библии";

        public const string PageNameDefaultBookOverview = "Общий обзор";
        public const string PageNameNotes = "Заметки";
        public const string PageNameDefaultComments = "Комментарии";

        public static readonly List<string> NotBibleStudyNotebooks = new List<string>() { "Личная", "Руководство по OneNote 2010" };
    }
}
