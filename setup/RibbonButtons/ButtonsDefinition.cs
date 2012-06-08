using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Extensibility;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.OneNote;
using Microsoft.Office.Core;
using System.Windows.Forms;
using System.Runtime.InteropServices.ComTypes;
using System.IO;
using System.Drawing.Imaging;
using System.Diagnostics;
using System.Reflection;
using System.Threading;
using System.Xml;
using System.Xml.Linq;

namespace RibbonButtons
{
    [GuidAttribute("61139959-A5E4-4261-977A-6262429033E1"), ProgId("RibbonButtons.ButtonsDefinition")]
	public class ButtonsDefinition : IDTExtensibility2, IRibbonExtensibility
	{
		#region IDTExtensibility2 Members

		ApplicationClass onApp = new ApplicationClass();

		public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
		{
			/*
				For debugging, it is useful to have a MessageBox.Show() here, so that execution is paused while you have a chance to get VS to 'Attach to Process' 
			*/
			onApp = (ApplicationClass)Application;
		}
		public void OnDisconnection(Extensibility.ext_DisconnectMode disconnectMode, ref System.Array custom)
		{
			//Clean up. Application is closing
			onApp = null;
			GC.Collect();
			GC.WaitForPendingFinalizers();

            MakeAllBibleReadOnly();
		}
		public void OnBeginShutdown(ref System.Array custom)
		{
			if (onApp != null)
				onApp = null;
		}
		public void OnStartupComplete(ref Array custom) { }
		public void OnAddInsUpdate(ref Array custom) { }

        public void MakeAllBibleReadOnly()
        {
            try
            {
                //todo: get the bible folder path from config file!
                foreach (var filePath in Directory.GetFiles(@"C:\Users\ademko\Documents\OneNote Notebooks\Библия", "*", SearchOption.AllDirectories))
                {
                    File.SetAttributes(filePath, FileAttributes.ReadOnly);
                }
            }
            catch (Exception ex)
            {
                //todo: log it
            }
        }

		#endregion

		#region IRibbonExtensibility Members

		/// <summary>
		/// Called at the start of the running of the add-in. Loads the ribbon
		/// </summary>
		public string GetCustomUI(string RibbonID)
		{   
			return Properties.Resources.ribbon;
		}

		/// <summary>
		/// Called from the onAction="" parameter in ribbon.xml. This is effectivley the onClick() function
		/// </summary>
		/// <param name="control">The control that was just clicked. control.Id will give you its ID</param>
        public void ButtonClick(IRibbonControl control)
		{
            string path = null;
            string args = string.Empty;                

            switch (control.Id)
            {
                case "VersePointerButton":
                    //UnLockCurrentSection();
                    //return;
                    path = Path.Combine(Utils.GetCurrentDirectory(), "tools\\BibleVersePointer\\BibleVersePointer.exe");
                    break;
                case "VerseLinkerButton":
                    path = Path.Combine(Utils.GetCurrentDirectory(), "tools\\BibleVerseLinker\\BibleVerseLinkerEx.exe");
                    break;
                case "NoteLinkerButton":
                    path = Path.Combine(Utils.GetCurrentDirectory(), "tools\\BibleNoteLinker\\BibleNoteLinker.exe");
                    break;
                case "ConfigureButton":
                    path = Path.Combine(Utils.GetCurrentDirectory(), "tools\\BibleConfigurator\\BibleConfigurator.exe");
                    break;
                case "HelpButton":
                    path = Path.Combine(Utils.GetCurrentDirectory(), "tools\\BibleConfigurator\\BibleConfigurator.exe");
                    args = "-showManual";
                    break;
                case "ModuleInfoButton":
                    path = Path.Combine(Utils.GetCurrentDirectory(), "tools\\BibleConfigurator\\BibleConfigurator.exe");
                    args = "-showModuleInfo";
                    break;
                case "AboutProgramButton":
                    path = Path.Combine(Utils.GetCurrentDirectory(), "tools\\BibleConfigurator\\BibleConfigurator.exe");
                    args = "-showAboutProgram";
                    break;
                    
            }
             

            if (!string.IsNullOrEmpty(path))
                Process.Start(path, args);
		}

        private void SyncCurrentPage()
        {
            onApp.SyncHierarchy(onApp.Windows.CurrentWindow.CurrentPageId);
        }

        private void UnLockCurrentSection()
        {
            //app.SyncHierarchy(app.Windows.CurrentWindow.CurrentSectionId);            

            string xml;
            onApp.GetHierarchy(onApp.Windows.CurrentWindow.CurrentSectionId, Microsoft.Office.Interop.OneNote.HierarchyScope.hsSelf, out xml);
            XmlNamespaceManager xnm;
            var xDoc = GetXDocument(xml, out xnm);
            string path = xDoc.Root.Attribute("path").Value;
            FileAttributes attribs = File.GetAttributes(path);


            if (attribs != FileAttributes.ReadOnly)
            {
                new Thread(SyncCurrentPage).Start();
                Thread.Sleep(1000);
                File.SetAttributes(path, FileAttributes.ReadOnly);
            }
            else
            {
                File.SetAttributes(path, FileAttributes.Normal);
                new Thread(SyncCurrentPage).Start();
            }

            //app.SyncHierarchy(app.Windows.CurrentWindow.CurrentPageId);
            //app.SyncHierarchy(app.Windows.CurrentWindow.CurrentSectionId);

            //            MessageBox.Show("All done! 8");
        }

        public static XDocument GetXDocument(string xml, out XmlNamespaceManager xnm)
        {
            XDocument xd = XDocument.Parse(xml);
            xnm = new XmlNamespaceManager(new NameTable());
            xnm.AddNamespace("one", "http://schemas.microsoft.com/office/onenote/2010/onenote");
            return xd;
        }

		/// <summary>
		/// Called from the loadImage="" parameter in ribbon.xml. Converts the images into IStreams
		/// </summary>
		/// <param name="imageName">The image="" parameter in ribbon.xml, i.e. the image name</param>
        public IStream GetImage(string imageName)
        {
            MemoryStream mem = new MemoryStream();

            switch (imageName)
            {
                case "VersePointerButton.png":
                    Properties.Resources.VersePointerButton.Save(mem, ImageFormat.Png);
                    break;
                case "VerseLinkerButton.png":
                    Properties.Resources.VerseLinkerButton.Save(mem, ImageFormat.Png);
                    break;
                case "NoteLinkerButton.png":
                    Properties.Resources.NoteLinkerButton.Save(mem, ImageFormat.Png);
                    break;
                case "ConfigureButton.png":
                    Properties.Resources.ConfigureButton.Save(mem, ImageFormat.Png);
                    break;
                case "HelpButton.png":
                    Properties.Resources.HelpButton.Save(mem, ImageFormat.Png);
                    break;
                case "AboutModule.png":
                    Properties.Resources.AboutModule.Save(mem, ImageFormat.Png);
                    break;
                case "AboutProgram.png":
                    Properties.Resources.AboutProgram.Save(mem, ImageFormat.Png);
                    break;
            }

            return new CCOMStreamWrapper(mem);

        }

		#endregion
	}
}