﻿using System;
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
        public static readonly string NewToolsName = "BibleNote";
        public static readonly string ConfigFileName = "settings.config";
        public static readonly string ModulesDirectoryName = "Modules";
        public static readonly string ModulesPackagesDirectoryName = "ModulesPackages";
        public static readonly string ManifestFileName = "manifest.xml";
        public static readonly string BibleContentFileName = "bible.xml";
        public static readonly string DictionaryContentsFileName = "dictionary.xml";
        public static readonly string FileExtensionIsbt = ".isbt";
        public static readonly string FileExtensionOnepkg = ".onepkg";
        public static readonly string FileExtensionXml = ".xml";
        public static readonly string FileExtensionCache = ".cache";
        public static readonly string TempDirectory = "Temp";
        public static readonly string CacheDirectory = "Cache";
        public static readonly string NotesPagesDirectory = "NotesPages";
        public static readonly string AnalyzedVersesDirectory = "AnalyzedVerses";        
        public static readonly string LogsDirectory = "Logs";
        public static readonly string DefaultPartVersesAlphabet = "abcdefghijklmnopqrstuvwxyz";
        public static readonly int DefaultStrongNumbersCount = 14700;
        public static readonly string UnicodeFontName = "Arial Unicode MS";
        public static readonly string AnalysisMutix = "ISBT_Analysis";
        public static readonly string ParametersMutix = "ISBT_Parameters";
        public static readonly string OneNoteProtocol = "onenote:";
        public static readonly string DoNotAnalyzeSymbol1 = "{}";
        public static readonly string DoNotAnalyzeSymbol2 = "[]";
        public static readonly string NotEmptyVerseContentSymbol = "[]";
        public static readonly string NoLinkTransmitHref = "#nt";

        public static readonly decimal ImportantVerseWeight = 2;

        public static readonly Version ModulesWithXmlBibleMinVersion = new Version(1, 9);

        public static readonly int DefaultPageWidth_Notes = 500;
        public static readonly int DefaultPageWidth_RubbishNotes = 500;
        public static readonly int DefaultPageWidth_Bible = 500;
        public static readonly bool Default_UseProxyLinksForStrong = true;
        public static readonly bool Default_UseProxyLinksForBibleVerses = true;
        public static readonly bool DefaultExpandMultiVersesLinking = true;
        public static readonly bool DefaultExcludedVersesLinking = false;
        public static readonly bool DefaultUseDifferentPagesForEachVerse = true;
        public static readonly bool DefaultRubbishPage_Use = false;
        public static readonly bool DefaultRubbishPage_ExpandMultiVersesLinking = true;
        public static readonly bool DefaultRubbishPage_ExcludedVersesLinking = true;
        public static readonly decimal DefaultFilter_MinVerseWeight = 0;
        public static readonly bool DefaultFilter_ShowDetailedNotes = false;

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

        public static readonly string ParameterName_FolderPathBibleNotesPages = "FolderPath_BibleNotesPages";

        public static readonly string ParameterName_ModuleName = "ModuleName";
        public static readonly string ParameterName_SupplementalBibleModules = "SupplementalBibleModules";
        public static readonly string ParameterName_SupplementalBibleLinkName = "SupplementalBibleLinkName";
        public static readonly string ParameterName_DictionariesModules = "DictionariesModules";
        public static readonly string ParameterName_SelectedNotebooksForAnalyze = "SelectedNotebooksForAnalyze";
        public static readonly string ParameterName_ShownMessages = "ShownMessages";

        public static readonly string ParameterName_FilterHiddenNotebooks = "FilterHiddenNotebooks";
        public static readonly string ParameterName_FilterMinVerseWeight = "FilterMinVerseWeight";
        public static readonly string ParameterName_FilterShowDetailedNotes = "FilterShowDetailedNotes";

        public static readonly string ParameterName_VersionFromSettings = "ProgramVersion";        


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

        public static readonly string ParameterName_UseProxyLinksForStrong = "UseProxyLinksForStrong";
        public static readonly string ParameterName_UseProxyLinksForBibleVerses = "UseProxyLinksForBibleVerses";
        public static readonly string ParameterName_UseProxyLinksForLinks = "UseProxyLinksForLinks";

        public static readonly string ParameterName_UseDefaultSettings = "UseDefaultSettings";

        public static readonly string ParameterName_Language = "Language";

        public static readonly string ParameterName_GenerateFullBibleVersesCache = "GenerateFullBibleVersesCache";


        public static readonly TimeSpan NewVersionCheckPeriod = new TimeSpan(1, 0, 0, 0);

        public static readonly string QueryParameter_QuickAnalyze = "qa=1";
        public static readonly string QueryParameter_BibleVerse = "bv=1";
        public static readonly string QueryParameter_ExtendedVerse = "ev=1";
        public static readonly string QueryParameterKey_VersePosition = "vp";
        public static readonly string QueryParameterKey_VerseWeight = "vw";
        public static readonly string QueryParameterKey_NotePageId = "npid";
        public static readonly string QueryParameterKey_CustomPageId = "cpId";
        public static readonly string QueryParameterKey_CustomObjectId = "coId";
        public static readonly string QueryParameterKey_IsDetailedLink = "idl";

        public static readonly string Key_IsSummaryNotesPage = "IsSummaryNotesPage";
        public static readonly string Key_LatestAnalyzeTime = "LatestAnalyzeTime";
        public static readonly string Key_EmbeddedBibleModule = "BibleModule";
        public static readonly string Key_EmbeddedSupplementalModules = "SupplementalModules";
        public static readonly string Key_EmbeddedDictionaries = "Dictionaries";
        public static readonly string Key_NotesPageManagerName = "NotesPageManagerName";
        public static readonly string Key_Id = "ID";
        public static readonly string Key_Verse = "Verse";
        public static readonly string Key_NotesPageLink = "NotesPageLink";
        public static readonly string Key_SyncId = "SyncID";
        public static readonly string Key_Table = "isbtTable";


        public static readonly int ChapterNotesPageLinkOutline_OffsetX = 41;
        public static readonly int ChapterNotesPageLinkOutline_y = 45;
        public static readonly int ChapterNotesPageLinkOutline_z = 1;

        public static readonly string NotesPageElementAttributeName_SyncId = "syncid";
        public static readonly string NotesPageStyleFileName = "core.css";
        public static readonly string NotesPageScriptFileName = "core.js";
        public static readonly string NotesPageJQueryScriptFileName = "jquery.js";
        public static readonly string NotesPageTrackbarStyleFileName = "plugins/trackbar/trackbar.css";
        public static readonly string NotesPageTrackbarScriptFileName = "plugins/trackbar/trackbar.js";
    }
}
