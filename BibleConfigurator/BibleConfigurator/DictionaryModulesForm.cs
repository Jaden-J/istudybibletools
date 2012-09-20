using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace BibleConfigurator
{
    public class DictionaryModulesForm: BaseSupplementalForm
    {
        public DictionaryModulesForm(Microsoft.Office.Interop.OneNote.Application oneNoteApp, MainForm form)
            : base(oneNoteApp, form)
        { }

        protected override string GetValidSupplementalNotebookId()
        {
            throw new NotImplementedException();
        }

        protected override void DeleteSupplementalModulesInSettingsStorage()
        {
            throw new NotImplementedException();
        }

        protected override int GetSupplementalModulesCount()
        {
            throw new NotImplementedException();
        }

        protected override bool SupplementalModuleAlreadyAdded(string moduleShortName)
        {
            throw new NotImplementedException();
        }

        protected override string FormDescription
        {
            get { throw new NotImplementedException(); }
        }

        protected override List<string> CommitChanges(BibleCommon.Common.ModuleInfo selectedModuleInfo)
        {
            throw new NotImplementedException();
        }

        protected override string GetSupplementalModuleName(int index)
        {
            throw new NotImplementedException();
        }

        protected override bool CanModuleBeDeleted(int index)
        {
            throw new NotImplementedException();
        }

        protected override void DeleteModule(string moduleShortName)
        {
            throw new NotImplementedException();
        }

        protected override string CloseSupplementalNotebookConfirmText
        {
            get { throw new NotImplementedException(); }
        }

        protected override void CloseSupplementalNotebook()
        {
            throw new NotImplementedException();
        }

        protected override bool IsModuleSupported(BibleCommon.Common.ModuleInfo moduleInfo)
        {
            throw new NotImplementedException();
        }
    }
}
