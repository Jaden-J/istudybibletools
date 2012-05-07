using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using BibleCommon.Common;
using BibleConfigurator.ModuleConverter;
using System.Text;
using BibleCommon.Services;


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


            //var converter = new BibleQuotaConverter("Test", @"C:\BibleQuote\RCCV", @"c:\manifest.xml", Encoding.Unicode,
            //    "1. Old Testament", "2. New Testament", 39, 27, new List<NotebookInfo>() 
            //    {
            //        new NotebookInfo() { Type = NotebookType.Single, Name = "Holy Bible.onepkg", SectionGroups = new List<BibleCommon.Common.SectionGroupInfo>()
            //        {
            //            new BibleCommon.Common.SectionGroupInfo() { Type = SectionGroupType.Bible, Name="Bible" },
            //            new BibleCommon.Common.SectionGroupInfo() { Type = SectionGroupType.BibleStudy, Name="Bible Study" },
            //            new BibleCommon.Common.SectionGroupInfo() { Type = SectionGroupType.BibleComments, Name="Bible Comments" }                        
            //        } },
            //        new NotebookInfo() { Type = NotebookType.Bible, Name = "Bible.onepkg" },
            //        new NotebookInfo() { Type = NotebookType.BibleStudy, Name = "Bible Study.onepkg" },
            //        new NotebookInfo() { Type = NotebookType.BibleComments, Name = "Notes on the Bible.onepkg" },
            //        new NotebookInfo() { Type = NotebookType.BibleNotesPages, Name = "Summary Notes.onepkg" }
            //    });
            //converter.Convert();

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            Form form;
            if (args.Length > 0
                    && args[0] == Consts.ShowModuleInfo
                    && SettingsManager.Instance.IsConfigured(new Microsoft.Office.Interop.OneNote.Application()))
                form = new AboutModuleForm(SettingsManager.Instance.ModuleName, true);
            else
                form = new MainForm(args);

            Application.Run(form);
        }

      
    }
}
