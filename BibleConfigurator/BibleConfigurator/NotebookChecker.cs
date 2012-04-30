﻿using System;
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

namespace BibleConfigurator
{
    public static class NotebookChecker
    {
        public static bool CheckNotebook(Application oneNoteApp, ModuleInfo module, string notebookId, NotebookType notebookType)
        {
            //string errorText = string.Empty;
            bool result = false;

            if (!string.IsNullOrEmpty(notebookId))
            {
                OneNoteProxy.HierarchyElement notebook = OneNoteProxy.Instance.GetHierarchy(oneNoteApp, notebookId, HierarchyScope.hsSections, true);                

                switch (notebookType)
                {
                    case NotebookType.Single:
                        result = ElementIsSingleNotebook(notebook.Content, notebook.Xnm);
                        break;
                    case NotebookType.Bible:
                        result = ElementIsBible(notebook.Content.Root, notebook.Xnm);
                        break;
                    case NotebookType.BibleComments:
                    case NotebookType.BibleNotesPages:
                        result = ElementIsBibleComments(notebook.Content.Root, notebook.Xnm);
                        break;
                    case NotebookType.BibleStudy:
                        result = ElementIsBibleStudy(module, notebook.Content.Root, notebook.Xnm);
                        break;
                }
            }

            return result;
        }

        public static bool ElementIsBibleStudy(ModuleInfo module, XElement element, XmlNamespaceManager xnm)
        {
            bool result = !(ElementIsBible(element, xnm) || ElementIsBibleComments(element, xnm));
            

            //if (result)
            //{
            //    //string notebookName = (string)element.Attribute("name").Value;
            //    //if (Consts.NotBibleStudyNotebooks.Contains(notebookName))                
            //    //    result = false;

            //    //string bibleStudyNotebookDefaultName =
            //    //    Path.GetFileNameWithoutExtension(module.Notebooks.First(n => n.Type == NotebookType.BibleStudy).Name);

            //    //if (!notebookName.StartsWith(bibleStudyNotebookDefaultName))
            //    //    result = false;
            //}


            return result;
        }

        public static bool ElementIsSingleNotebook(XDocument notebookDoc, XmlNamespaceManager xnm)
        {
            List<XElement> sectionsGroups = notebookDoc.Root.XPathSelectElements("one:SectionGroup", xnm).Where(sg => !OneNoteUtils.IsRecycleBin(sg)).ToList();

            if (sectionsGroups.Count == 3)
            {
                if ((ElementIsBible(sectionsGroups[0], xnm) || ElementIsBible(sectionsGroups[1], xnm) || ElementIsBible(sectionsGroups[2], xnm))
                    && (ElementIsBibleComments(sectionsGroups[0], xnm) || ElementIsBibleComments(sectionsGroups[1], xnm) || ElementIsBibleComments(sectionsGroups[2], xnm)))
                    return true;
            }

            return false;
        }

        public static bool ElementIsBible(XElement element, XmlNamespaceManager xnm)
        {
            //todo: переделать, когда будет поддержка модулей
            XElement oldTestamentSectionGroup = element.XPathSelectElement(string.Format("one:SectionGroup[@name='{0}']", Consts.OldTestamentName), xnm);

            if (oldTestamentSectionGroup != null)
            {
                int oldTestamentSectionsCount = oldTestamentSectionGroup.XPathSelectElements("one:Section", xnm).Count();

                if (oldTestamentSectionsCount > 35)  
                {
                    XElement newTestamentSectionGroup = element.XPathSelectElement(string.Format("one:SectionGroup[@name='{0}']", Consts.NewTestamentName), xnm);

                    if (newTestamentSectionGroup != null)
                    {
                        int newTestamentSectionsCount = newTestamentSectionGroup.XPathSelectElements("one:Section", xnm).Count();

                        if (newTestamentSectionsCount > 25)
                        {
                            return true;
                        }
                    }
                }
            }

            return false;
        }

        public static bool ElementIsBibleComments(XElement element, XmlNamespaceManager xnm)
        {
            //todo: переделать, когда будет поддержка модулей
            XElement oldTestamentSectionGroup = element.XPathSelectElement(string.Format("one:SectionGroup[@name='{0}']", Consts.OldTestamentName), xnm);

            if (oldTestamentSectionGroup != null)
            {
                int subSectionsCount = oldTestamentSectionGroup.XPathSelectElements("one:Section", xnm).Count();

                if (subSectionsCount < 5)
                {
                    XElement newTestamentSectionGroup = element.XPathSelectElement(string.Format("one:SectionGroup[@name='{0}']", Consts.NewTestamentName), xnm);

                    if (newTestamentSectionGroup != null)
                    {
                        subSectionsCount = newTestamentSectionGroup.XPathSelectElements("one:Section", xnm).Count();

                        if (subSectionsCount  < 5)
                        {
                            return true;
                        }
                    }
                }
            }

            return false;
        }
    }
}
