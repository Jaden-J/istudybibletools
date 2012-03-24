using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Configuration.Install;
using System.IO;
using System.Diagnostics;
using System.Reflection;
using System.Collections;
using BibleCommon.Helpers;

namespace BibleConfigurator
{

    // Taken from:http://msdn2.microsoft.com/en-us/library/
    // system.configuration.configurationmanager.aspx
    // Set 'RunInstaller' attribute to true.

    [RunInstaller(true)]
    public class Installer : System.Configuration.Install.Installer
    {
        public Installer()
            : base()
        {
            // Attach the 'Committed' event.
            this.Committed += new InstallEventHandler(MyInstaller_Committed);
            // Attach the 'Committing' event.
            this.Committing += new InstallEventHandler(MyInstaller_Committing);
        }

        // Event handler for 'Committing' event.
        private void MyInstaller_Committing(object sender, InstallEventArgs e)
        {
            //Console.WriteLine("");
            //Console.WriteLine("Committing Event occurred.");
            //Console.WriteLine("");
        }

        // Event handler for 'Committed' event.
        private void MyInstaller_Committed(object sender, InstallEventArgs e)
        {
            //Directory.SetCurrentDirectory(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location));
            //Process.Start(Path.Combine(Utils.GetCurrentDirectory(), "BibleConfigurator.exe"), "-RunAfterSetup");


            string currentDirectory = Utils.GetCurrentDirectory();
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.UseShellExecute = true;
            startInfo.WorkingDirectory = currentDirectory;
            startInfo.FileName = Path.Combine(currentDirectory, "BibleConfigurator.exe");
            startInfo.Verb = "runas";
            startInfo.Arguments = "-RunAfterSetup";

            Process p = Process.Start(startInfo);
        }

        // Override the 'Install' method.
        public override void Install(IDictionary savedState)
        {
            base.Install(savedState);
        }

        // Override the 'Commit' method.
        public override void Commit(IDictionary savedState)
        {
            base.Commit(savedState);
        }

        // Override the 'Rollback' method.
        public override void Rollback(IDictionary savedState)
        {
            base.Rollback(savedState);
        }
    }
}
