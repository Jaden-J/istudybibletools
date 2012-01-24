using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.OneNote;

namespace BibleConfigurator.tools
{
    public class UpdateCommentLinksManager
    {
        private Application _oneNoteApp;

        public void UpdateLinks(Application oneNoteApp)
        {
            _oneNoteApp = oneNoteApp;

            BibleCommon.Services.Logger.Init("UpdateCommentLinksManager");

            BibleCommon.Services.Logger.Done();
        }

        private static void ProcessNotebook(Application oneNoteApp, string notebookId, string sectionGroupId, Args userArgs)
        {
            BibleCommon.Services.Logger.LogMessage("Обработка записной книжки: '{0}'", OneNoteUtils.GetHierarchyElementName(oneNoteApp, notebookId));  // чтобы точно убедиться

            string hierarchyXml;
            oneNoteApp.GetHierarchy(notebookId, HierarchyScope.hsPages, out hierarchyXml);
            XmlNamespaceManager xnm;
            XDocument notebookDoc = OneNoteUtils.GetXDocument(hierarchyXml, out xnm);

            Logger.MoveLevel(1);
            ProcessRootSectionGroup(oneNoteApp, notebookId, notebookDoc, sectionGroupId, xnm, userArgs.AnalyzeDepth, userArgs.Force, userArgs.DeleteNotes);
            Logger.MoveLevel(-1);
        }
    }
}
