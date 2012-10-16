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

namespace BibleCommon.Services
{
    public static class DictionaryTermsCacheManager
    {
        private static string GetCacheFilePath(string moduleShortName)
        {
            var dictionaryModuleInfo = SettingsManager.Instance.DictionariesModules.FirstOrDefault(m => m.ModuleName == moduleShortName);
            if (dictionaryModuleInfo == null)
                throw new ArgumentException("The dictionary '{0}' is not installed", moduleShortName);

            return Path.Combine(Utils.GetCacheFolderPath(), dictionaryModuleInfo.ToString()) + ".cache";
        }

        public static bool CacheIsActive(string moduleShortName)
        {
            return File.Exists(GetCacheFilePath(moduleShortName));
        }

        public static DictionaryCachedTermSet LoadCachedDictionary(string moduleShortName)
        {
            string filePath = GetCacheFilePath(moduleShortName);
            if (!File.Exists(filePath))
                throw new NotConfiguredException(string.Format("The cache file of '{0}' does not exist: '{1}'", moduleShortName, filePath));

            return (DictionaryCachedTermSet)BinarySerializerHelper.Deserialize(filePath);
        }

        public static void GenerateCache(Application oneNoteApp, ModuleInfo moduleInfo, ICustomLogger logger)
        {
            var cacheData = IndexDictionary(oneNoteApp, moduleInfo, logger);
            var filePath = GetCacheFilePath(moduleInfo.ShortName);

            BinarySerializerHelper.Serialize(cacheData, filePath);
        }

        public static DictionaryCachedTermSet IndexDictionary(Application oneNoteApp, ModuleInfo moduleInfo, ICustomLogger logger)
        {
            var result = new DictionaryCachedTermSet();
            var dictionaryModuleInfo = SettingsManager.Instance.DictionariesModules.FirstOrDefault(m => m.ModuleName == moduleInfo.ShortName);
            if (dictionaryModuleInfo != null)
            {
                XmlNamespaceManager xnm;
                var sectionGroupDoc = OneNoteUtils.GetHierarchyElement(oneNoteApp, dictionaryModuleInfo.SectionId, HierarchyScope.hsPages, out xnm);

                var sectionsEl = sectionGroupDoc.Root.XPathSelectElements("one:Section", xnm);
                if (sectionsEl.Count() > 0)
                {
                    foreach (var sectionEl in sectionsEl)
                    {
                        IndexDictionarySection(oneNoteApp, sectionEl, result, moduleInfo.Type == ModuleType.Strong, logger, xnm);
                    }
                }
                else
                    IndexDictionarySection(oneNoteApp, sectionGroupDoc.Root, result, moduleInfo.Type == ModuleType.Strong, logger, xnm);
            }

            return result;
        }

        private static void IndexDictionarySection(Application oneNoteApp, XElement sectionEl, DictionaryCachedTermSet result, bool isStrong, ICustomLogger logger, XmlNamespaceManager xnm)
        {
            string sectionName = (string)sectionEl.Attribute("name");

            foreach (var pageEl in sectionEl.XPathSelectElements("one:Page", xnm))
            {
                var pageId = (string)pageEl.Attribute("ID");
                var pageDoc = OneNoteUtils.GetPageContent(oneNoteApp, pageId, out xnm);

                var tableEl = NotebookGenerator.GetPageTable(pageDoc, xnm);

                foreach (var termTextEl in tableEl.XPathSelectElements("one:Row/one:Cell[1]/one:OEChildren/one:OE/one:T", xnm))
                {
                    var termName = StringUtils.GetText(termTextEl.Value);
                    termName = isStrong ? termName.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)[0] : termName.ToLower();
                    var termTextElementId = (string)termTextEl.Parent.Attribute("objectID");

                    var href = isStrong ? OneNoteProxy.Instance.GenerateHref(oneNoteApp, pageId, termTextElementId) : null;                    

                    if (!result.ContainsKey(termName))
                        result.Add(termName, new DictionaryTermLink() { PageId = pageId, ObjectId = termTextElementId, Href = href });

                    if (logger != null)
                        logger.LogMessage(termName);
                }               
            }
        }
    }
}
