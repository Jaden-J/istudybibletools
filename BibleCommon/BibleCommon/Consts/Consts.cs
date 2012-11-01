using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections.Specialized;
using Microsoft.Office.Interop.OneNote;
using System.Xml.Linq;

namespace BibleCommon.Consts
{
    public static class Constants
    {
        public static readonly string OneNoteXmlNs = "http://schemas.microsoft.com/office/onenote/2010/onenote";
        public static readonly XMLSchema CurrentOneNoteSchema = XMLSchema.xs2010;
        public static readonly string ToolsName = "IStudyBibleTools";
        public static readonly string ConfigFileName = "settings.config";
        public static readonly string ModulesDirectoryName = "Modules";
        public static readonly string ModulesPackagesDirectoryName = "ModulesPackages";
        public static readonly string ManifestFileName = "manifest.xml";
        public static readonly string BibleInfoFileName = "bible.xml";
        public static readonly string DictionaryInfoFileName = "dictionary.xml";
        public static readonly string FileExtensionIsbt = ".isbt";
        public static readonly string FileExtensionOnepkg = ".onepkg";
        public static readonly string TempDirectory = "Temp";
        public static readonly string CacheDirectory = "Cache";        
        public static readonly string DefaultPartVersesAlphabet = "abcdefghijklmnopqrstuvwxyz";
        public static readonly int DefaultStrongNumbersCount = 14700;
        
        public static readonly int DefaultPageWidth_Notes = 500;
        public static readonly int DefaultPageWidth_RubbishNotes = 500;
        public static readonly int DefaultPageWidth_Bible = 500;
        public static readonly bool DefaultExpandMultiVersesLinking = true;
        public static readonly bool DefaultExcludedVersesLinking = false;
        public static readonly bool DefaultUseDifferentPagesForEachVerse = true;
        public static readonly bool DefaultRubbishPage_Use = false;
        public static readonly bool DefaultRubbishPage_ExpandMultiVersesLinking = true;
        public static readonly bool DefaultRubbishPage_ExcludedVersesLinking = true;        
        public static readonly bool DefaultUseMiddleStrongLinks = true;

        public static readonly string ResourceName_DefaultPageNameDefaultBookOverview = "DefaultPageNameDefaultBookOverview";
        public static readonly string ResourceName_DefaultPageNameDefaultComments = "DefaultPageNameDefaultComments";
        public static readonly string ResourceName_DefaultPageName_Notes = "DefaultPageName_Notes";
        public static readonly string ResourceName_DefaultPageName_RubbishNotes = "DefaultPageName_RubbishNotes";
        public static readonly string ResourceName_DefaultSupplementalBibleLinkName = "DefaultSupplementalBibleLinkName";        

        public static readonly string ParameterName_NotebookIdBible = "NotebookId_Bible";
        public static readonly string ParameterName_NotebookIdBibleComments = "NotebookId_BibleComments";
        public static readonly string ParameterName_NotebookIdBibleNotesPages = "NotebookId_BibleNotesPages";
        public static readonly string ParameterName_NotebookIdBibleStudy = "NotebookId_BibleStudy";
        public static readonly string ParameterName_NotebookIdSupplementalBible = "NotebookId_SupplementalBible";
        public static readonly string ParameterName_NotebookIdDictionaries = "NotebookId_Dictionaries";
        public static readonly string ParameterName_SectionGroupIdBible = "SectionGroupId_Bible";        
        public static readonly string ParameterName_SectionGroupIdBibleStudy = "SectionGroupId_BibleStudy";
        public static readonly string ParameterName_SectionGroupIdBibleComments = "SectionGroupId_BibleComments";
        public static readonly string ParameterName_SectionGroupIdBibleNotesPages = "SectionGroupId_BibleNotesPages";
        public static readonly string ParameterName_PageNameDefaultComments = "PageName_DefaultComments";
        public static readonly string ParameterName_SectionNameDefaultBookOverview = "SectionName_DefaultBookOverview";        
        public static readonly string ParameterName_PageNameNotes = "PageName_Notes";
        public static readonly string ParameterName_LastNotesLinkTime = "LastNotesLinkTime";
        public static readonly string ParameterName_NewVersionOnServer = "NewVersionOnServer";
        public static readonly string ParameterName_NewVersionOnServerLatestCheckTime = "NewVersionOnServerLatestCheckTime";

        public static readonly string ParameterName_ModuleName = "ModuleName";
        public static readonly string ParameterName_SupplementalBibleModules = "SupplementalBibleModules";
        public static readonly string ParameterName_SupplementalBibleLinkName = "SupplementalBibleLinkName";
        public static readonly string ParameterName_DictionariesModules = "DictionariesModules";
        

        public static readonly string ParameterName_PageWidthNotes = "Width_NotesPage";
        public static readonly string ParameterName_PageWidthBible = "Width_BiblePage";
        public static readonly string ParameterName_ExpandMultiVersesLinking = "ExpandMultiVersesLinking";
        public static readonly string ParameterName_ExcludedVersesLinking = "ExcludedVersesLinking";
        public static readonly string ParameterName_UseDifferentPagesForEachVerse = "UseDifferentPagesForEachVerse";
        public static readonly string ParameterName_RubbishPageUse = "RubbishPage_Use";
        public static readonly string ParameterName_PageNameRubbishNotes = "RubbishPage_NotesPageName";
        public static readonly string ParameterName_PageWidthRubbishNotes = "RubbishPage_NotesPageWidth";
        public static readonly string ParameterName_RubbishPageExpandMultiVersesLinking = "RubbishPage_ExpandMultiVersesLinking";
        public static readonly string ParameterName_RubbishPageExcludedVersesLinking = "RubbishPage_ExcludedVersesLinking";

        public static readonly string ParameterName_UseMiddleStrongLinks = "UseMiddleStrongLinks";

        public static readonly string ParameterName_UseDefaultSettings = "UseDefaultSettings";        

        public static readonly string ParameterName_Language = "Language";        
        
        public static readonly TimeSpan NewVersionCheckPeriod = new TimeSpan(1, 0, 0, 0);        
        

        public static readonly string Key_IsSummaryNotesPage = "IsSummaryNotesPage";
        public static readonly string Key_LatestAnalyzeTime = "LatestAnalyzeTime";        
    }   
}

