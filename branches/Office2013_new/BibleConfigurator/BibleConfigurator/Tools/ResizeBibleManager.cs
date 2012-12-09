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
    public class ResizeBibleManager: IDisposable
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
            if (!SettingsManager.Instance.IsConfigured(ref _oneNoteApp))
            {
                FormLogger.LogError(BibleCommon.Resources.Constants.Error_SystemIsNotConfigured);
                return;
            }   

            try
            {
                BibleCommon.Services.Logger.Init("ResizeBibleManager");

                try
                {
                    OneNoteLocker.UnlockBible(ref _oneNoteApp);
                }
                catch (NotSupportedException)
                {
                    //todo: log it
                }

                int chaptersCount = ModulesManager.GetBibleChaptersCount(SettingsManager.Instance.ModuleShortName, true);
                _form.PrepareForLongProcessing(chaptersCount, 1, BibleCommon.Resources.Constants.ResizeBibleTableManagerStartMessage);

                NotebookIteratorHelper.Iterate(ref _oneNoteApp,
                    SettingsManager.Instance.NotebookId_Bible, SettingsManager.Instance.SectionGroupId_Bible, pageInfo =>
                        {
                            try
                            {
                                ResizeBiblePage(pageInfo.Id, pageInfo.Title, width);
                            }
                            catch (Exception ex)
                            {
                                FormLogger.LogError(ex.ToString());
                            }

                            if (_form.StopLongProcess)
                                throw new ProcessAbortedByUserException();
                        });

                _form.LongProcessingDone(BibleCommon.Resources.Constants.ResizeBibleTableManagerFinishMessage);                
            }
            catch (ProcessAbortedByUserException)
            {
                BibleCommon.Services.Logger.LogMessageParams("Process aborted by user");
            }
            finally
            {                
                BibleCommon.Services.Logger.Done();                
            }
        }      

        private void ResizeBiblePage(string pageId, string pageName, int width)
        {
            _form.PerformProgressStep(string.Format("{0} '{1}'", BibleCommon.Resources.Constants.ProcessPage, pageName));

            string pageContentXml = null;
            XDocument notePageDocument;
            XmlNamespaceManager xnm;
            OneNoteUtils.UseOneNoteAPI(ref _oneNoteApp, () =>
            {
                _oneNoteApp.GetPageContent(pageId, out pageContentXml, PageInfo.piBasic, Constants.CurrentOneNoteSchema);
            });
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

                OneNoteUtils.UpdatePageContentSafe(ref _oneNoteApp, notePageDocument, xnm);
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
            column.SetAttributeValue("width", width);
        }

        public void Dispose()
        {
            _oneNoteApp = null;
            _form = null;
        }
    }
}
