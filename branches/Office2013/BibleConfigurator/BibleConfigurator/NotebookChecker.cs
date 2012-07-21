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
        public static bool CheckNotebook(Application oneNoteApp, ModuleInfo module, string notebookId, NotebookType notebookType, out string errorText)
        {
            errorText = string.Empty;

            if (!string.IsNullOrEmpty(notebookId))
            {
                OneNoteProxy.HierarchyElement notebook = OneNoteProxy.Instance.GetHierarchy(oneNoteApp, notebookId, HierarchyScope.hsSections, true);

                try
                {
                    switch (notebookType)
                    {
                        case NotebookType.Single:
                            CheckElementIsSingleNotebook(module, notebook.Content, notebook.Xnm);
                            break;
                        case NotebookType.Bible:
                            CheckElementIsBible(module, notebook.Content.Root, notebook.Xnm);
                            break;
                        case NotebookType.BibleComments:
                        case NotebookType.BibleNotesPages:
                            CheckElementIsBibleComments(module, notebook.Content.Root, notebook.Xnm);
                            break;
                        case NotebookType.BibleStudy:
                            CheckElementIsBibleStudy(module, notebook.Content.Root, notebook.Xnm);
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

        private static void CheckElementIsBibleStudy(ModuleInfo module, XElement element, XmlNamespaceManager xnm)
        {
            if (ElementIsBible(module, element, xnm))
                throw new InvalidNotebookException(Constants.SelectedNotebookForType, NotebookType.Bible);

            if (ElementIsBibleComments(module, element, xnm))
                throw new InvalidNotebookException(Constants.SelectedNotebookForType, NotebookType.BibleComments);            
        }

        private static void CheckElementIsSingleNotebook(ModuleInfo module, XDocument notebookDoc, XmlNamespaceManager xnm)
        {
            List<XElement> sectionsGroups = notebookDoc.Root.XPathSelectElements("one:SectionGroup", xnm).Where(sg => !OneNoteUtils.IsRecycleBin(sg)).ToList();

            if (sectionsGroups.Count != 3)
                throw new InvalidNotebookException(Constants.WrongSectionGroupsCount, 3, sectionsGroups.Count);            
            
            if (!(ElementIsBible(module, sectionsGroups[0], xnm) || ElementIsBible(module, sectionsGroups[1], xnm) || ElementIsBible(module, sectionsGroups[2], xnm)))
                throw new InvalidNotebookException(Constants.SectionGroupOfTypeNotFound, SectionGroupType.Bible);            


            if (!(ElementIsBibleComments(module, sectionsGroups[0], xnm) || ElementIsBibleComments(module, sectionsGroups[1], xnm) || ElementIsBibleComments(module, sectionsGroups[2], xnm)))
                throw new InvalidNotebookException(Constants.SectionGroupOfTypeNotFound, SectionGroupType.BibleComments);                               
        }

        public static bool ElementIsBible(ModuleInfo module, XElement element, XmlNamespaceManager xnm)
        {
            try
            {
                CheckElementIsBible(module, element, xnm);
                return true;
            }
            catch (InvalidNotebookException)
            {
                return false;
            }

        }

        private static void CheckElementIsBible(ModuleInfo module, XElement element, XmlNamespaceManager xnm)
        {            
            XElement oldTestamentSectionGroup = element.XPathSelectElement(string.Format("one:SectionGroup[@name='{0}']", module.BibleStructure.OldTestamentName), xnm);

            if (oldTestamentSectionGroup == null)
                throw new InvalidNotebookException(Constants.SectionGroupNotFound, module.BibleStructure.OldTestamentName);

            int oldTestamentSectionsCount = oldTestamentSectionGroup.XPathSelectElements("one:Section", xnm).Count();

            if (oldTestamentSectionsCount < module.BibleStructure.OldTestamentBooksCount)
                throw new InvalidNotebookException(Constants.WrongSectionsCountInSectionGroup,
                    module.BibleStructure.OldTestamentName, module.BibleStructure.OldTestamentBooksCount, oldTestamentSectionsCount);

            XElement newTestamentSectionGroup = element.XPathSelectElement(string.Format("one:SectionGroup[@name='{0}']", module.BibleStructure.NewTestamentName), xnm);

            if (newTestamentSectionGroup == null)
                throw new InvalidNotebookException(Constants.SectionGroupNotFound, module.BibleStructure.NewTestamentName);

            int newTestamentSectionsCount = newTestamentSectionGroup.XPathSelectElements("one:Section", xnm).Count();

            if (newTestamentSectionsCount < module.BibleStructure.NewTestamentBooksCount)
                throw new InvalidNotebookException(Constants.WrongSectionsCountInSectionGroup,
                    module.BibleStructure.NewTestamentName, module.BibleStructure.NewTestamentBooksCount, newTestamentSectionsCount);
        }

        public static bool ElementIsBibleComments(ModuleInfo module, XElement element, XmlNamespaceManager xnm)
        {
            try
            {
                CheckElementIsBibleComments(module, element, xnm);
                return true;
            }
            catch (InvalidNotebookException)
            {
                return false;
            }

        }

        private static void CheckElementIsBibleComments(ModuleInfo module, XElement element, XmlNamespaceManager xnm)
        {
            XElement oldTestamentSectionGroup = element.XPathSelectElement(string.Format("one:SectionGroup[@name='{0}']", module.BibleStructure.OldTestamentName), xnm);

            if (oldTestamentSectionGroup == null)
                throw new InvalidNotebookException(Constants.SectionGroupNotFound, module.BibleStructure.OldTestamentName);

            int subSectionsCount = oldTestamentSectionGroup.XPathSelectElements("one:Section", xnm).Count();

            if (subSectionsCount > 3)
                throw new InvalidNotebookException(Constants.WrongSectionsCountInSectionGroup, 0, module.BibleStructure.OldTestamentName, subSectionsCount);

            XElement newTestamentSectionGroup = element.XPathSelectElement(string.Format("one:SectionGroup[@name='{0}']", module.BibleStructure.NewTestamentName), xnm);

            if (newTestamentSectionGroup == null)
                throw new InvalidNotebookException(Constants.SectionGroupNotFound, module.BibleStructure.NewTestamentName);

            subSectionsCount = newTestamentSectionGroup.XPathSelectElements("one:Section", xnm).Count();

            if (subSectionsCount > 3)
                throw new InvalidNotebookException(Constants.WrongSectionsCountInSectionGroup, 0, module.BibleStructure.NewTestamentName, subSectionsCount);
        }

    }
}
