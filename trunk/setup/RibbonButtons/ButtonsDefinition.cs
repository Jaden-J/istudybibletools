﻿using System;
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
		}
		public void OnBeginShutdown(ref System.Array custom)
		{
			if (onApp != null)
				onApp = null;
		}
		public void OnStartupComplete(ref Array custom) { }
		public void OnAddInsUpdate(ref Array custom) { }

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

            switch (control.Id)
            {
                case "VersePointerButton":                                            
                    path = Path.Combine(Utils.GetApplicationFolder(), "tools\\BibleVersePointer\\BibleVersePointer.exe");
                    break;
                case "VerseLinkerButton":
                    path = Path.Combine(Utils.GetApplicationFolder(), "tools\\BibleVerseLinker\\BibleVerseLinkerEx.exe");
                    break;
                case "NoteLinkerButton":
                    path = Path.Combine(Utils.GetApplicationFolder(), "tools\\BibleNoteLinker\\BibleNoteLinkerEx.exe");
                    break;
                case "ConfigureButton":
                    path = Path.Combine(Utils.GetApplicationFolder(), "tools\\BibleConfigurator\\BibleConfigurator.exe");
                    break;
                case "HelpButton":
                    path = Path.Combine(Utils.GetApplicationFolder(), "Instruction.docx");
                    break;
            }
             

            if (!string.IsNullOrEmpty(path))
                Process.Start(path);
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
            }

            return new CCOMStreamWrapper(mem);

        }

		#endregion
	}
}