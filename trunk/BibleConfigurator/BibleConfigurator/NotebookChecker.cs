using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Xml;
using System.Xml.XPath;
using Microsoft.Office.Interop.OneNote;
using BibleCommon.Helpers;

namespace BibleConfigurator
{
    public static class NotebookChecker
    {
        public static bool CheckNotebook(Application oneNoteApp, string notebookId, NotebookType notebookType)
        {
            //string errorText = string.Empty;
            bool result = false;

            if (!string.IsNullOrEmpty(notebookId))
            {
                string xml;
                XmlNamespaceManager xnm;
                oneNoteApp.GetHierarchy(notebookId, HierarchyScope.hsSections, out xml);
                XDocument notebookDoc = OneNoteUtils.GetXDocument(xml, out xnm);

                switch (notebookType)
                {
                    case NotebookType.Single:
                        result = ElementIsSingleNotebook(notebookDoc, xnm);
                        break;
                    case NotebookType.Bible:
                        result = ElementIsBible(notebookDoc.Root, xnm);
                        break;
                    case NotebookType.BibleComments:
                        result = ElementIsBibleComments(notebookDoc.Root, xnm);
                        break;
                    case NotebookType.BibleStudy:
                        result = ElementIsBibleStudy(notebookDoc.Root, xnm);
                        break;
                }
            }

            return result;
        }

        public static bool ElementIsBibleStudy(XElement element, XmlNamespaceManager xnm)
        {
            bool result = !(ElementIsBible(element, xnm) || ElementIsBibleComments(element, xnm));
            

            if (result)
            {
                string notebookName = (string)element.Attribute("name").Value;
                //if (Consts.NotBibleStudyNotebooks.Contains(notebookName))                
                //    result = false;

                if (!notebookName.StartsWith(Consts.BibleStudyNotebookDefaultName))
                    result = false;
            }


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
            XElement oldTestamentSectionGroup = element.XPathSelectElement(string.Format("one:SectionGroup[@name='{0}']", Consts.OldTestamentName), xnm);

            if (oldTestamentSectionGroup != null)
            {
                int oldTestamentSectionsCount = oldTestamentSectionGroup.XPathSelectElements("one:Section", xnm).Count();

                if (oldTestamentSectionsCount == 39)
                {
                    XElement newTestamentSectionGroup = element.XPathSelectElement(string.Format("one:SectionGroup[@name='{0}']", Consts.NewTestamentName), xnm);

                    if (newTestamentSectionGroup != null)
                    {
                        int newTestamentSectionsCount = newTestamentSectionGroup.XPathSelectElements("one:Section", xnm).Count();

                        if (newTestamentSectionsCount == 27)
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
            XElement oldTestamentSectionGroup = element.XPathSelectElement(string.Format("one:SectionGroup[@name='{0}']", Consts.OldTestamentName), xnm);

            if (oldTestamentSectionGroup != null)
            {
                int subSectionsCount = oldTestamentSectionGroup.XPathSelectElements("one:Section", xnm).Count();

                if (subSectionsCount == 0)
                {
                    XElement newTestamentSectionGroup = element.XPathSelectElement(string.Format("one:SectionGroup[@name='{0}']", Consts.NewTestamentName), xnm);

                    if (newTestamentSectionGroup != null)
                    {
                        subSectionsCount = newTestamentSectionGroup.XPathSelectElements("one:Section", xnm).Count();

                        if (subSectionsCount == 0)
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
