using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using BibleCommon.Common;
using BibleConfigurator.ModuleConverter;
using System.Text;
using BibleCommon.Services;
using BibleCommon.Helpers;
using System.IO;
using System.Diagnostics;

namespace BibleConfigurator
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(params string[] args)
        {
            //DefaultRusModuleGenerator.GenerateModuleInfo();

            var converter = new BibleQuotaConverter("Test", @"G:\Dropbox\RCCV", @"c:\manifest.xml", Encoding.Unicode,
                "Vechiul Testament", "Noul Testament", 39, 27, new List<NotebookInfo>() 
            //var converter = new BibleQuotaConverter("Test", @"G:\Dropbox\Изучение Библии\программы\Цитата из Библии\King_James_Version", @"c:\manifest.xml", Encoding.ASCII,
            //    "1. Old Testament", "2. New Testament", 39, 27, new List<NotebookInfo>() 
                {
                    //new NotebookInfo() { Type = NotebookType.Single, Name = "Holy Bible.onepkg", SectionGroups = new List<BibleCommon.Common.SectionGroupInfo>()
                    //{
                    //    new BibleCommon.Common.SectionGroupInfo() { Type = SectionGroupType.Bible, Name="Bible" },
                    //    new BibleCommon.Common.SectionGroupInfo() { Type = SectionGroupType.BibleStudy, Name="Bible Study" },
                    //    new BibleCommon.Common.SectionGroupInfo() { Type = SectionGroupType.BibleComments, Name="Bible Comments" }                        
                    //} },
                    new NotebookInfo() { Type = NotebookType.Bible, Name = "Bible.onepkg" },
                    new NotebookInfo() { Type = NotebookType.BibleStudy, Name = "Bible Study.onepkg" },
                    new NotebookInfo() { Type = NotebookType.BibleComments, Name = "Comments to the Bible.onepkg" },
                    new NotebookInfo() { Type = NotebookType.BibleNotesPages, Name = "Summary of Notes.onepkg" }
                });
            converter.Convert();

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            Form form = null;
            if (args.Contains(Consts.ShowModuleInfo) && SettingsManager.Instance.IsConfigured(new Microsoft.Office.Interop.OneNote.Application()))
                form = new AboutModuleForm(SettingsManager.Instance.ModuleName, true);
            else if (args.Contains(Consts.ShowAboutProgram))
                form = new AboutProgramForm();
            else if (args.Contains(Consts.ShowManual))
                OpenManual();
            else if (args.Length == 1)
            {
                string moduleFilePath = args[0];
                if (File.Exists(moduleFilePath))
                {
                    form = new MainForm(args);
                    bool needToReload = ((MainForm)form).AddNewModule(moduleFilePath);
                    ((MainForm)form).ShowModulesTabAtStartUp = true;
                    ((MainForm)form).NeedToSaveChangesAfterLoadingModuleAtStartUp = needToReload;
                }
            }
            
            if (form == null)
                form = new MainForm(args);

            Application.Run(form);
        }

        private static void OpenManual()
        {
            string path = Path.GetDirectoryName(Path.GetDirectoryName(Utils.GetCurrentDirectory()));

            path = Path.Combine(path, "Instruction_" + LanguageManager.UserLanguage.LCID + ".htm");

            Process.Start(path);
        }

      
    }
}
