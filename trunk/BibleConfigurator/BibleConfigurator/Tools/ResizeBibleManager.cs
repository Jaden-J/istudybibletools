using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.OneNote;
using BibleCommon.Helpers;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;
using BibleCommon.Services;
using BibleCommon.Consts;
using BibleCommon.Common;

namespace BibleConfigurator.Tools
{
    public class ResizeBibleManager
    {
        private Application _oneNoteApp;
        private MainForm _form;

        public ResizeBibleManager(Application oneNoteApp, MainForm form)
        {
            _oneNoteApp = oneNoteApp;
            _form = form;
        }

        public void ResizeBiblePages(int width)
        {
            if (!SettingsManager.Instance.IsConfigured(_oneNoteApp))
            {
                Logger.LogError(BibleCommon.Resources.Constants.Error_SystemIsNotConfigures);
                return;
            }   

            try
            {
                BibleCommon.Services.Logger.Init("ResizeBibleManager");

                _form.PrepareForExternalProcessing(1255, 1, BibleCommon.Resources.Constants.ResizeBibleTableManagerStartMessage);

                NotebookIteratorHelper.Iterate(_oneNoteApp,
                    SettingsManager.Instance.NotebookId_Bible, SettingsManager.Instance.SectionGroupId_Bible, pageInfo =>
                        {
                            try
                            {
                                ResizeBiblePage(pageInfo.Id, pageInfo.Title, width);
                            }
                            catch (Exception ex)
                            {
                                Logger.LogError(ex.ToString());
                            }

                            if (_form.StopExternalProcess)
                                throw new ProcessAbortedByUserException();
                        });

                _form.ExternalProcessingDone(BibleCommon.Resources.Constants.ResizeBibleTableManagerFinishMessage);                
            }
            catch (ProcessAbortedByUserException)
            {
                BibleCommon.Services.Logger.LogMessage("Process aborted by user");
            }
            finally
            {                
                BibleCommon.Services.Logger.Done();                
            }
        }      

        private void ResizeBiblePage(string pageId, string pageName, int width)
        {
            _form.PerformProgressStep(string.Format("{0} '{1}'", BibleCommon.Resources.Constants.ProcessPage, pageName));

            string pageContentXml;
            XDocument notePageDocument;
            XmlNamespaceManager xnm;
            _oneNoteApp.GetPageContent(pageId, out pageContentXml);
            notePageDocument = OneNoteUtils.GetXDocument(pageContentXml, out xnm);

            XElement columns = notePageDocument.Root.XPathSelectElement("//one:Table/one:Columns", xnm);
            if (columns != null)
            {

                XElement column1 = columns.XPathSelectElement("one:Column[1]", xnm);
                XElement column2 = columns.XPathSelectElement("one:Column[2]", xnm);

                SetWidthAttribute(column1, width);
                SetLockedAttribute(column1);

                SetWidthAttribute(column2, 37);
                SetLockedAttribute(column2);

                OneNoteUtils.UpdatePageContentSafe(_oneNoteApp, notePageDocument);
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

        private static void SetWidthAttribute(XElement column, int width)
        {
            column.Attribute("width").Value = width.ToString();
        }
    }
}
