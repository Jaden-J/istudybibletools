using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using BibleCommon.Helpers;
using BibleCommon.Common;
using System.Xml;
using Microsoft.Office.Interop.OneNote;
using BibleCommon.Contracts;
using System.Xml.XPath;
using System.Xml.Linq;
using Polenter.Serialization;
using BibleCommon.UI.Forms;

namespace BibleCommon.Services
{
    public static class DictionaryTermsCacheManager
    {
        private static string GetCacheFilePath(string moduleShortName)
        {
            var dictionaryModuleInfo = SettingsManager.Instance.DictionariesModules.FirstOrDefault(m => m.ModuleName == moduleShortName);
            if (dictionaryModuleInfo == null)
                throw new Exception(string.Format("The dictionary '{0}' is not installed", moduleShortName));

            return Path.Combine(Utils.GetCacheFolderPath(), dictionaryModuleInfo.ToString()) + "_terms.cache";
        }

        public static bool CacheIsActive(string moduleShortName)
        {
            return File.Exists(GetCacheFilePath(moduleShortName));
        }

        public static Dictionary<string, string> LoadCachedDictionary(string moduleShortName)
        {
            string filePath = GetCacheFilePath(moduleShortName);
            if (!File.Exists(filePath))
                throw new NotConfiguredException(string.Format("The cache file of '{0}' does not exist: '{1}'", moduleShortName, filePath));

            return SharpSerializationHelper.Deserialize<Dictionary<string, string>>(filePath);
        }

        public static void RemoveCache(string moduleShortName)
        {
            try
            {
                var filePath = GetCacheFilePath(moduleShortName);
                if (File.Exists(filePath))
                    File.Delete(filePath);
            }
            catch (Exception ex)
            {
                BibleCommon.Services.Logger.LogError(ex);
            }
        }

        public static Dictionary<string, string> GenerateCache(ref Application oneNoteApp, ModuleInfo moduleInfo, ICustomLogger logger, out List<string> notFoundTerms)
        {
            notFoundTerms = null;

            var cacheData = IndexDictionary(ref oneNoteApp, moduleInfo, logger);

            var filePath = GetCacheFilePath(moduleInfo.ShortName);

            SharpSerializationHelper.Serialize(cacheData, filePath);

            if (moduleInfo.NotebooksStructure.DictionaryTermsCount > cacheData.Count)
            {
                int notFoundTermsCount;
                if (moduleInfo.Type == ModuleType.Dictionary)
                {
                    notFoundTerms = new List<string>();
                    bool isStrong = moduleInfo.Type == ModuleType.Strong;

                    var moduleDictionaryInfo = ModulesManager.GetModuleDictionaryInfo(moduleInfo.ShortName);
                    foreach (var term in moduleDictionaryInfo.TermSet.Terms)
                    {
                        if (!cacheData.ContainsKey(GetTermName(term, isStrong)))
                            notFoundTerms.Add(term);
                    }

                    notFoundTermsCount = notFoundTerms.Count;                    
                }
                else
                    notFoundTermsCount = moduleInfo.NotebooksStructure.DictionaryTermsCount.GetValueOrDefault() - cacheData.Count;

                Logger.LogMessageParams("{0} terms were not found in dictionary '{1}'", notFoundTermsCount, moduleInfo.ShortName);
            }

            return cacheData;
        }

        private static Dictionary<string, string> IndexDictionary(ref Application oneNoteApp, ModuleInfo moduleInfo, ICustomLogger logger)
        {
            var result = new Dictionary<string, string>();
            var dictionaryModuleInfo = SettingsManager.Instance.DictionariesModules.FirstOrDefault(m => m.ModuleName == moduleInfo.ShortName);
            if (dictionaryModuleInfo != null)
            {
                XmlNamespaceManager xnm;
                var sectionGroupDoc = OneNoteUtils.GetHierarchyElement(ref oneNoteApp, dictionaryModuleInfo.SectionId, HierarchyScope.hsPages, out xnm);

                var sectionsEl = sectionGroupDoc.Root.XPathSelectElements("one:Section", xnm);
                if (sectionsEl.Count() > 0)
                {
                    foreach (var sectionEl in sectionsEl)
                    {
                        IndexDictionarySection(ref oneNoteApp, sectionEl, ref result, moduleInfo.Type == ModuleType.Strong, logger, xnm);
                    }
                }
                else
                    IndexDictionarySection(ref oneNoteApp, sectionGroupDoc.Root, ref result, moduleInfo.Type == ModuleType.Strong, logger, xnm);
            }           

            return result;
        }

        private static void IndexDictionarySection(ref Application oneNoteApp, XElement sectionEl, ref Dictionary<string, string> result, bool isStrong, ICustomLogger logger, XmlNamespaceManager xnm)
        {
            string sectionName = (string)sectionEl.Attribute("name");

            foreach (var pageEl in sectionEl.XPathSelectElements("one:Page", xnm))
            {
                var pageId = (string)pageEl.Attribute("ID");
                var pageDoc = OneNoteUtils.GetPageContent(ref oneNoteApp, pageId, out xnm);

                var tableEl = NotebookGenerator.GetPageTable(pageDoc, xnm);

                foreach (var termTextEl in tableEl.XPathSelectElements("one:Row/one:Cell[1]/one:OEChildren/one:OE/one:T", xnm))
                {
                    var termName = StringUtils.GetText(termTextEl.Value);
                    if (isStrong)
                        termName = termName.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)[0];
                    termName = GetTermName(termName, isStrong);

                    var termTextElementId = (string)termTextEl.Parent.Attribute("objectID");

                    var href = (isStrong && !SettingsManager.Instance.UseProxyLinksForStrong) 
                                    ? ApplicationCache.Instance.GenerateHref(ref oneNoteApp, pageId, termTextElementId) 
                                    : null;                    

                    if (!result.ContainsKey(termName))
                        result.Add(termName, new DictionaryTermLink() { PageId = pageId, ObjectId = termTextElementId, Href = href }.ToString());

                    if (logger != null)
                        logger.LogMessage(termName);
                }               
            }
        }        

        private static string GetTermName(string term, bool isStrong)
        {
            return isStrong ? term : term.ToLower();
        }
    }
}
