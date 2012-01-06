using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Xml.XPath;
using Microsoft.Office.Interop.OneNote;
using BibleCommon;
using System.Xml;
using BibleCommon.Services;
using BibleCommon.Helpers;

namespace BibleTablesResizer
{
    public static class TableModifier
    {
        public static void ModifyTable(Application oneNoteApp, string pageId)
        {
            try
            {                
                string pageContentXml;
                XDocument notePageDocument;
                XmlNamespaceManager xnm;
                oneNoteApp.GetPageContent(pageId, out pageContentXml);
                notePageDocument = Utils.GetXDocument(pageContentXml, out xnm);

                XElement columns = notePageDocument.Root.XPathSelectElement("//one:Table/one:Columns", xnm);
                if (columns != null)
                {

                    XElement column1 = columns.XPathSelectElement("one:Column[1]", xnm);
                    XElement column2 = columns.XPathSelectElement("one:Column[2]", xnm);

                    SetWidthAttribute(column1, "500");
                    SetLockedAttribute(column1);

                    SetWidthAttribute(column2, "37");
                    SetLockedAttribute(column2);  
                 
                    oneNoteApp.UpdatePageContent(notePageDocument.ToString());
                }
            }
            catch (Exception ex)
            {
                Logger.LogError("Ошибки при обработке страницы.", ex);
            }
        }

        private static void SetLockedAttribute(XElement column)
        {
            XAttribute isLocked = column.Attribute("isLocked");
            if (isLocked != null)
                isLocked.Value = "true";                    
            else
                column.Add(new XAttribute("isLocked", true));
        }

        private static void SetWidthAttribute(XElement column, string width)
        {
            column.Attribute("width").Value = width;                    
        }
    }
}
