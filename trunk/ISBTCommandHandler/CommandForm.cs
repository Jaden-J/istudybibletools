using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using BibleCommon.Contracts;
using BibleCommon.Handlers;
using System.Threading;
using BibleCommon.Services;
using BibleCommon.Common;
using BibleCommon.Helpers;

namespace ISBTCommandHandler
{
    public partial class CommandForm : Form
    {
        private IProtocolHandler[] _handlers = new IProtocolHandler[] 
                                                    { 
                                                        new QuickAnalyzeHandler(), 
                                                        new OpenBibleVerseHandler(),
                                                        new OpenNotesPageHandler(),
                                                        new NavigateToStrongHandler(), 
                                                        new FindVersesWithStrongNumberHandler(),
                                                        new RefreshCacheHandler(),
                                                        new ExitApplicationHandler()
                                                    };

        public CommandForm()
        {
            InitializeComponent();

            var args = Environment.GetCommandLineArgs();
            if (args.Length > 0)
                ProcessCommandLine(args);
        }

        protected override CreateParams CreateParams
        {
            get
            {
                var cp = base.CreateParams;
                cp.ExStyle |= 0x80;  // Turn on WS_EX_TOOLWINDOW
                return cp;                
            }
        }

        protected override bool ShowWithoutActivation
        {
            get { return true; }
        }

        internal void ProcessCommandLine(params string[] args)
        {
            try
            {
                if (args.Length > 1)
                    args = args.ToList().Skip(1).ToArray();

                foreach (var handler in _handlers)
                {
                    if (handler.IsProtocolCommand(args))
                    {
                        if (handler is ExitApplicationHandler)
                        {
                            if (_notesPageForm != null)
                                _notesPageForm.ExitApplication = true;
                        }

                        handler.ExecuteCommand(args);

                        if (handler is OpenNotesPageHandler)
                        {
                            OpenNotesPage(((OpenNotesPageHandler)handler).Verse, ((OpenNotesPageHandler)handler).GetVerseFilePath());
                        }
                        else if (handler is RefreshCacheHandler)
                        {
                            if (((RefreshCacheHandler)handler).CacheMode == RefreshCacheHandler.RefreshCacheMode.RefreshAnalyzedVersesCache)
                            {
                                if (_notesPageForm != null)
                                    _notesPageForm.RefreshFilteredNotebooksInfo();
                            }
                        }

                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                FormLogger.LogError(ex);
            }
        }

        private NotesPageForm _notesPageForm = null;
        private void OpenNotesPage(VersePointer vp, string filePath)
        {
            try
            {
                if (_notesPageForm == null)
                {
                    _notesPageForm = new NotesPageForm();
                    _notesPageForm.ShowInTaskbar = true;
                }

                this.SetFocus();
                _notesPageForm.OpenNotesPage(vp, filePath);
            }
            catch (Exception ex)
            {
                FormLogger.LogError(ex);
            }            
        }               
    }
}
