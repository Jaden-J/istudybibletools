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

        private static void CheckContainer(SectionGroupInfo container, XElement containerEl, XmlNamespaceManager xnm)
        {
            if (container.SkipCheck)
                return;

            if (container.CheckSectionGroupsCount)
            {
                int subSectionGroupsCount = containerEl.XPathSelectElements("one:SectionGroup", xnm).Count();

                if (container.SectionGroupsCount != default(int))
                {
                    if (container.SectionGroupsCount != subSectionGroupsCount)
                        throw new InvalidNotebookException(Constants.SectionGroupsCountNotEqual, container.Name, container.SectionGroupsCount, subSectionGroupsCount);
                }
                else if (container.SectionGroupsCountMin != default(int))
                {
                    if (container.SectionGroupsCountMin > subSectionGroupsCount)
                        throw new InvalidNotebookException(Constants.SectionGroupsCountLessThanMin, container.Name, container.SectionGroupsCountMin, subSectionGroupsCount);
                }
                else if (container.SectionGroupsCountMax != default(int))
                {
                    if (container.SectionGroupsCountMax < subSectionGroupsCount)
                        throw new InvalidNotebookException(Constants.SectionGroupsCountMoreThanMax, container.Name, container.SectionGroupsCountMax, subSectionGroupsCount);
                }
            }

            if (container.CheckSectionsCount)
            {
                int sectionsCount = containerEl.XPathSelectElements("one:Section", xnm).Count();

                if (container.SectionsCount != default(int))
                {
                    if (container.SectionsCount != sectionsCount)
                        throw new InvalidNotebookException(Constants.SectionsCountNotEqual, container.Name, container.SectionsCount, sectionsCount);
                }
                else if (container.SectionsCountMin != default(int))
                {
                    if (container.SectionsCountMin > sectionsCount)
                        throw new InvalidNotebookException(Constants.SectionsCountLessThanMin, container.Name, container.SectionsCountMin, sectionsCount);
                }
                else if (container.SectionsCountMax != default(int))
                {
                    if (container.SectionsCountMax < sectionsCount)
                        throw new InvalidNotebookException(Constants.SectionsCountMoreThanMax, container.Name, container.SectionsCountMax, sectionsCount);
                }
            }

            foreach (var section in container.Sections)
            {
                var sectionEl = containerEl.XPathSelectElement(string.Format("one:Section", section.Name));

                if (section == null)
                    throw new InvalidNotebookException(Constants.SectionNotFoundInContainer, section.Name, container.Name);

                CheckSection(section, sectionEl, xnm);
            }

            foreach (var subSectionGroup in container.SectionGroups)
            {
                var subSectionGroupEl = containerEl.XPathSelectElement(string.Format("one:SectionGroup[@name='{0}']", subSectionGroup.Name), xnm);

                if (subSectionGroupEl == null)
                    throw new InvalidNotebookException(Constants.SectionGroupNotFoundInContainer, subSectionGroup.Name, container.Name);

                CheckContainer(subSectionGroup, subSectionGroupEl, xnm);
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

    }
}
