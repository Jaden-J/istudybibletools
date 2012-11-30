using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Xml;
using System.Xml.XPath;
using Microsoft.Office.Interop.OneNote;
using BibleCommon.Helpers;
using BibleCommon.Services;
using BibleCommon.Common;
using System.IO;
using BibleCommon.Resources;

namespace BibleConfigurator
{
    public static class NotebookChecker
    {
        public static bool CheckNotebook(Application oneNoteApp, ModuleInfo module, string notebookId, ContainerType notebookType, out string errorText)
        {
            errorText = string.Empty;

            if (!string.IsNullOrEmpty(notebookId))
            {
                OneNoteProxy.HierarchyElement notebookEl = OneNoteProxy.Instance.GetHierarchy(oneNoteApp, notebookId, HierarchyScope.hsSections, true);
                var notebook = module.GetNotebook(notebookType);

                try
                {
                    switch (notebookType)
                    {   
                        case ContainerType.BibleStudy:
                            CheckElementIsBibleStudy(module, notebookEl.Content.Root, notebookEl.Xnm);
                            break;
                        default:
                            CheckContainer(notebook, notebookEl.Content.Root, notebookEl.Xnm);
                            CheckNotebookMetadata(oneNoteApp, module, notebookId, notebookType, notebookEl);     // то есть проверка показала, что записная книжка похожа на Библию. Теперь посмотрим, есть ли информация в метаданных                       
                            break;
                    }

                    return true;
                }
                catch (InvalidNotebookException ex)
                {
                    errorText = ex.Message;
                }
            }

            return false;
        }

        public static bool ElementIsBible(ModuleInfo module, XElement element, XmlNamespaceManager xnm)
        {
            try
            {
                CheckContainer(module.GetNotebook(ContainerType.Bible), element, xnm);
                return true;
            }
            catch (InvalidNotebookException)
            {
                return false;
            }

        }

        public static bool ElementIsBibleComments(ModuleInfo module, XElement element, XmlNamespaceManager xnm)
        {
            try
            {
                CheckContainer(module.GetNotebook(ContainerType.BibleComments), element, xnm);
                return true;
            }
            catch (InvalidNotebookException)
            {
                return false;
            }
        }     

        internal static XElement GetFirstNotebookPageId(Application oneNoteApp, string notebookId, OneNoteProxy.HierarchyElement containerEl, out XmlNamespaceManager xnm)
        {   
            XElement sectionsDoc;

            if (containerEl == null)
                sectionsDoc = OneNoteUtils.GetHierarchyElement(oneNoteApp, notebookId, HierarchyScope.hsSections, out xnm).Root;
            else
            {
                sectionsDoc = containerEl.Content.Root;
                xnm = containerEl.Xnm;
            }

            var firstSection = sectionsDoc.XPathSelectElement(string.Format("//one:Section[{0}]", OneNoteUtils.NotInRecycleXPathCondition), xnm);
            if (firstSection != null)
            {
                var pagesDoc = OneNoteUtils.GetHierarchyElement(oneNoteApp, (string)firstSection.Attribute("ID"), HierarchyScope.hsPages, out xnm);
                var firstPage = pagesDoc.Root.XPathSelectElement("one:Page", xnm);
                return firstPage;
            }

            return null;
        }

        private static void CheckNotebookMetadata(Application oneNoteApp, ModuleInfo module, string notebookId, ContainerType notebookType, OneNoteProxy.HierarchyElement containerEl)
        {
            if (notebookType == ContainerType.Bible)
            {
                XmlNamespaceManager xnm;
                var firstNotebookPageEl = GetFirstNotebookPageId(oneNoteApp, notebookId, containerEl, out xnm);
                if (firstNotebookPageEl != null)
                {
                    var bibleModuleMetadata = OneNoteUtils.GetPageMetaData(oneNoteApp, firstNotebookPageEl, BibleCommon.Consts.Constants.Key_EmbeddedBibleModule, xnm);
                    if (!string.IsNullOrEmpty(bibleModuleMetadata))
                    {
                        var bibleModuleInfo = EmbeddedModuleInfo.Deserialize(bibleModuleMetadata);
                        if (bibleModuleInfo.Count > 0)
                        {
                            if (bibleModuleInfo[0].ModuleName != module.ShortName)
                            {
                                var containerName = (string)containerEl.Content.Root.Attribute("name");
                                throw new InvalidNotebookException(BibleCommon.Resources.Constants.BibleNotebookIsForAnotherModule,
                                                                        containerName, bibleModuleInfo[0].ModuleName, module.ShortName);
                            }
                        }
                    }
                }
            }
        }

        private static void CheckContainer(SectionGroupInfo container, XElement containerEl, XmlNamespaceManager xnm)
        {
            if (container.SkipCheck)
                return;

            var containerName = (string)containerEl.Attribute("name");

            if (container.CheckSectionGroupsCount)
            {
                int subSectionGroupsCount = containerEl.XPathSelectElements("one:SectionGroup", xnm).Count();

                if (container.SectionGroupsCount != default(int))
                {
                    if (container.SectionGroupsCount != subSectionGroupsCount)
                        throw new InvalidNotebookException(Constants.SectionGroupsCountNotEqual, containerName, container.SectionGroupsCount, subSectionGroupsCount);
                }
                else if (container.SectionGroupsCountMin != default(int))
                {
                    if (container.SectionGroupsCountMin > subSectionGroupsCount)
                        throw new InvalidNotebookException(Constants.SectionGroupsCountLessThanMin, containerName, container.SectionGroupsCountMin, subSectionGroupsCount);
                }
                else if (container.SectionGroupsCountMax != default(int))
                {
                    if (container.SectionGroupsCountMax < subSectionGroupsCount)
                        throw new InvalidNotebookException(Constants.SectionGroupsCountMoreThanMax, containerName, container.SectionGroupsCountMax, subSectionGroupsCount);
                }
            }

            if (container.CheckSectionsCount)
            {
                int sectionsCount = containerEl.XPathSelectElements("one:Section", xnm).Count();

                if (container.SectionsCount != default(int))
                {
                    if (container.SectionsCount != sectionsCount)
                        throw new InvalidNotebookException(Constants.SectionsCountNotEqual, containerName, container.SectionsCount, sectionsCount);
                }
                else if (container.SectionsCountMin != default(int))
                {
                    if (container.SectionsCountMin > sectionsCount)
                        throw new InvalidNotebookException(Constants.SectionsCountLessThanMin, containerName, container.SectionsCountMin, sectionsCount);
                }
                else if (container.SectionsCountMax != default(int))
                {
                    if (container.SectionsCountMax < sectionsCount)
                        throw new InvalidNotebookException(Constants.SectionsCountMoreThanMax, containerName, container.SectionsCountMax, sectionsCount);
                }
            }

            if (container.Sections != null)
            {
                foreach (var section in container.Sections)
                {
                    var sectionEl = containerEl.XPathSelectElement(string.Format("one:Section", section.Name));

                    if (section == null)
                        throw new InvalidNotebookException(Constants.SectionNotFoundInContainer, section.Name, containerName);

                    CheckSection(section, sectionEl, xnm);
                }
            }

            if (container.SectionGroups != null)
            {
                foreach (var subSectionGroup in container.SectionGroups)
                {
                    var subSectionGroupEl = containerEl.XPathSelectElement(string.Format("one:SectionGroup[@name='{0}']", subSectionGroup.Name), xnm);

                    if (subSectionGroupEl == null)
                        throw new InvalidNotebookException(Constants.SectionGroupNotFoundInContainer, subSectionGroup.Name, containerName);

                    CheckContainer(subSectionGroup, subSectionGroupEl, xnm);
                }
            }
        }

        private static void CheckSection(SectionInfo section, XElement sectionEl, XmlNamespaceManager xnm)
        {
            if (section.SkipCheck)
                return;

            int sectionsCount = sectionEl.XPathSelectElements("one:Page", xnm).Count();

            if (section.PagesCount != default(int))
            {
                if (section.PagesCount != sectionsCount)
                    throw new InvalidNotebookException(Constants.PagesCountNotEqual, section.Name, section.PagesCount, sectionsCount);
            }
            else if (section.PagesCountMin != default(int))
            {
                if (section.PagesCountMin > sectionsCount)
                    throw new InvalidNotebookException(Constants.PagesCountLessThanMin, section.Name, section.PagesCountMin, sectionsCount);
            }
            else if (section.PagesCountMax != default(int))
            {
                if (section.PagesCountMax < sectionsCount)
                    throw new InvalidNotebookException(Constants.PagesCountMoreThanMax, section.Name, section.PagesCountMax, sectionsCount);
            }
        }

        private static void CheckElementIsBibleStudy(ModuleInfo module, XElement element, XmlNamespaceManager xnm)
        {
            if (ElementIsBible(module, element, xnm))
                throw new InvalidNotebookException(Constants.SelectedNotebookForType, ContainerType.Bible);

            if (ElementIsBibleComments(module, element, xnm))
                throw new InvalidNotebookException(Constants.SelectedNotebookForType, ContainerType.BibleComments);            
        }      

    }
}
