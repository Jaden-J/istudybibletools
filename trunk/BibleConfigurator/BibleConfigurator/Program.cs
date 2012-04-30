using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using BibleCommon.Common;
using BibleConfigurator.ModuleConverter;
using System.Text;

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
            //ModuleGenerator.GenerateModuleInfo();


            //var converter = new BibleQuotaConverter("Test", @"G:\Dropbox\Изучение Библии\программы\Цитата из Библии\King_James_Version", @"c:\manifest.xml", Encoding.Default,
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
            //        new NotebookInfo() { Type = NotebookType.BibleComments, Name = "Bible Comments.onepkg" },
            //        new NotebookInfo() { Type = NotebookType.BibleNotesPages, Name = "Notes Pages.onepkg" }
            //    });
            //converter.Convert();

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm(args));
        }

      
    }
}
