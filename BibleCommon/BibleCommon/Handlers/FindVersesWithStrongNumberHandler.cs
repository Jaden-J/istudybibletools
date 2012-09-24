using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BibleCommon.Contracts;
using Microsoft.Office.Interop.OneNote;
using BibleCommon.Services;

namespace BibleCommon.Handlers
{
    public class FindVersesWithStrongNumberHandler : IProtocolHandler
    {
        public string GetCommandUrl(string strongNumber)
        {
            return string.Format("isbt_fvwsn://{0}", strongNumber);
        }

        public void ExecuteCommand(string strongNumber)
        {
            if (!string.IsNullOrEmpty(SettingsManager.Instance.NotebookId_SupplementalBible))
            {
                Application oneNoteApp = new Application();
                string result;
                oneNoteApp.FindPages(SettingsManager.Instance.NotebookId_SupplementalBible, strongNumber, out result, true, true, Consts.Constants.CurrentOneNoteSchema);
            }
        }
    }
}
