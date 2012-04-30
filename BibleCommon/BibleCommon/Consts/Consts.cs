using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections.Specialized;

namespace BibleCommon.Consts
{
    public static class Constants
    {
        public static readonly string OneNoteXmlNs = "http://schemas.microsoft.com/office/onenote/2010/onenote";
        public static readonly string ToolsName = "IStudyBibleTools";
        public static readonly string ConfigFileName = "settings.config";
        public static readonly string ModulesDirectoryName = "Modules";
        public static readonly string ModulesPackagesDirectoryName = "ModulesPackages";
        public static readonly string ManifestFileName = "manifest.xml";
        public static readonly string IsbtFileExtension = ".isbt";        
        public const string TempDirectory = "Temp";

        
        public static readonly int DefaultPageWidth_Notes = 500;
        public static readonly int DefaultPageWidth_RubbishNotes = 500;

        public static readonly string ParameterName_NotebookIdBible = "NotebookId_Bible";
        public static readonly string ParameterName_NotebookIdBibleComments = "NotebookId_BibleComments";
        public static readonly string ParameterName_NotebookIdBibleNotesPages = "NotebookId_BibleNotesPages";
        public static readonly string ParameterName_NotebookIdBibleStudy = "NotebookId_BibleStudy";        
        public static readonly string ParameterName_SectionGroupIdBible = "SectionGroupId_Bible";        
        public static readonly string ParameterName_SectionGroupIdBibleStudy = "SectionGroupId_BibleStudy";
        public static readonly string ParameterName_SectionGroupIdBibleComments = "SectionGroupId_BibleComments";
        public static readonly string ParameterName_SectionGroupIdBibleNotesPages = "SectionGroupId_BibleNotesPages";
        public static readonly string ParameterName_PageNameDefaultComments = "PageName_DefaultComments";
        public static readonly string ParameterName_PageNameDefaultBookOverview = "PageName_DefaultBookOverview";
        public static readonly string ParameterName_PageNameNotes = "PageName_Notes";
        public static readonly string ParameterName_LastNotesLinkTime = "LastNotesLinkTime";
        public static readonly string ParameterName_NewVersionOnServer = "NewVersionOnServer";
        public static readonly string ParameterName_NewVersionOnServerLatestCheckTime = "NewVersionOnServerLatestCheckTime";

        public static readonly string ParameterName_ModuleName = "ModuleName";

        public static readonly string ParameterName_PageWidthNotes = "Width_NotesPage";
        public static readonly string ParameterName_ExpandMultiVersesLinking = "ExpandMultiVersesLinking";
        public static readonly string ParameterName_ExcludedVersesLinking = "ExcludedVersesLinking";
        public static readonly string ParameterName_UseDifferentPagesForEachVerse = "UseDifferentPagesForEachVerse";
        public static readonly string ParameterName_RubbishPageUse = "RubbishPage_Use";
        public static readonly string ParameterName_PageNameRubbishNotes = "RubbishPage_NotesPageName";
        public static readonly string ParameterName_PageWidthRubbishNotes = "RubbishPage_NotesPageWidth";
        public static readonly string ParameterName_RubbishPageExpandMultiVersesLinking = "RubbishPage_ExpandMultiVersesLinking";
        public static readonly string ParameterName_RubbishPageExcludedVersesLinking = "RubbishPage_ExcludedVersesLinking";

        public static readonly string ParameterName_Language = "Language";

        public static readonly string NewVersionOnServerFileUrl = "http://IStudyBibleTools.ru/ServerVariables.xml";
        public static readonly string NewVersionOnServerCheckFileUrl = "http://IStudyBibleTools.ru/ServerVariables.htm";
        public static readonly TimeSpan NewVersionCheckPeriod = new TimeSpan(1, 0, 0, 0);

        public static readonly string DownloadPageUrl = "http://IStudyBibleTools.ru/download.htm?fromProgram=true";               
        

        public static readonly string Key_IsSummaryNotesPage = "IsSummaryNotesPage";
        public static readonly string Key_LatestAnalyzeTime = "LatestAnalyzeTime";
    }
}
