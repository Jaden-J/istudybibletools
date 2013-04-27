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
using System.Threading;
using System.Xml;
using System.Xml.Linq;

namespace RibbonButtons
{
    [GuidAttribute("61139959-A5E4-4261-977A-6262429033E1"), ProgId("IStudyBibleTools.ButtonsDefinition")]
	public class ButtonsDefinition : IDTExtensibility2, IRibbonExtensibility
    {
        #region consts

        private const string BibleConfiguratorPath = "tools\\BibleConfigurator\\BibleConfigurator.exe";
        private const string BibleCommonPath = "tools\\BibleConfigurator\\BibleCommon.dll";
        private const string SharpSerializerPath = "tools\\BibleConfigurator\\Polenter.SharpSerializer.dll";
        private const string BibleNoteLinkerPath = "tools\\BibleNoteLinker\\BibleNoteLinker.exe";
        private const string BibleVerseLinkerPath = "tools\\BibleVerseLinker\\BibleVerseLinkerEx.exe";
        private const string BibleVersePointerPath = "tools\\BibleVersePointer\\BibleVersePointer.exe";

        private const string BibleConfiguratorProgramClassName = "BibleConfigurator.Program";
        private const string BibleNoteLinkerProgramClassName = "BibleNoteLinker.Program";
        private const string BibleVerseLinkerProgramClassName = "BibleVerseLinkerEx.Program";
        private const string BibleVersePointerProgramClassName = "BibleVersePointer.Program";

        #endregion

        #region IDTExtensibility2 Members

        //ApplicationClass onApp;

		public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
		{
			/*
				For debugging, it is useful to have a MessageBox.Show() here, so that execution is paused while you have a chance to get VS to 'Attach to Process' 
			*/		

            try
            {
                //onApp = (ApplicationClass)Application;

                AppDomain.CurrentDomain.AssemblyResolve += new ResolveEventHandler(CurrentDomain_AssemblyResolve);
                RunProgram(Path.Combine(Utils.GetCurrentDirectory(), BibleConfiguratorPath), BibleConfiguratorProgramClassName, "-runOnOneNoteStarts", false);

                RunProgram("isbtRefreshCache:refreshCache", null, null, false);  // инициализируем кэш
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
		}

        Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            try
            {
                if (args.Name.Contains("BibleCommon, Version="))
                    return AssemblyLoader.LoadAssembly(Path.Combine(Utils.GetCurrentDirectory(), BibleCommonPath));
                else if (args.Name.Contains("Polenter.SharpSerializer, Version="))
                    return AssemblyLoader.LoadAssembly(Path.Combine(Utils.GetCurrentDirectory(), SharpSerializerPath));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            return null;
        }  

		public void OnDisconnection(Extensibility.ext_DisconnectMode disconnectMode, ref System.Array custom)
		{
            RunProgram("isbtExitApplication:exit", null, null, false);  // закрываем кэш
			//Clean up. Application is closing
			//onApp = null;
			GC.Collect();
			GC.WaitForPendingFinalizers();            
		}
		public void OnBeginShutdown(ref System.Array custom)
		{
            //if (onApp != null)
            //    onApp = null;
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
            try
            {
			    return Properties.Resources.ribbon;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return null;
            }
		}

		/// <summary>
		/// Called from the onAction="" parameter in ribbon.xml. This is effectivley the onClick() function
		/// </summary>
		/// <param name="control">The control that was just clicked. control.Id will give you its ID</param>
        public void ButtonClick(IRibbonControl control)
		{
            try
            {
                string path = null;
                string args = string.Empty;
                string programClassName = string.Empty;
                bool loadInSameProcess = false;

                switch (control.Id)
                {
                    case "VersePointerButton":
                        path = Path.Combine(Utils.GetCurrentDirectory(), BibleVersePointerPath);
                        programClassName = BibleVersePointerProgramClassName;
                        loadInSameProcess = true;
                        break;
                    case "VerseLinkerButton":
                        path = Path.Combine(Utils.GetCurrentDirectory(), BibleVerseLinkerPath);
                        programClassName = BibleVerseLinkerProgramClassName;
                        break;
                    case "NoteLinkerButton":
                        path = Path.Combine(Utils.GetCurrentDirectory(), BibleNoteLinkerPath);
                        programClassName = BibleNoteLinkerProgramClassName;
                        break;
                    case "QuickNoteLinkerButton":
                        path = Path.Combine(Utils.GetCurrentDirectory(), BibleNoteLinkerPath);
                        args = "-quickAnalyze";
                        programClassName = BibleNoteLinkerProgramClassName;
                        loadInSameProcess = true;
                        //path = "isbtQuickAnalyze:currentPage";
                        break;
                    case "ConfigureButton":
                        path = Path.Combine(Utils.GetCurrentDirectory(), BibleConfiguratorPath);
                        programClassName = BibleConfiguratorProgramClassName;
                        break;
                    case "HelpButton":
                        path = Path.Combine(Utils.GetCurrentDirectory(), BibleConfiguratorPath);
                        args = "-showManual";
                        programClassName = BibleConfiguratorProgramClassName;
                        //loadInSameProcess = true;
                        break;
                    case "ModuleInfoButton":
                        path = Path.Combine(Utils.GetCurrentDirectory(), BibleConfiguratorPath);
                        args = "-showModuleInfo";
                        programClassName = BibleConfiguratorProgramClassName;
                        break;
                    case "AboutProgramButton":
                        path = Path.Combine(Utils.GetCurrentDirectory(), BibleConfiguratorPath);
                        args = "-showAboutProgram";
                        programClassName = BibleConfiguratorProgramClassName;
                        break;
                    case "UnlockCurrentSection":
                        path = Path.Combine(Utils.GetCurrentDirectory(), BibleConfiguratorPath);
                        args = "-unlockBibleSection";
                        programClassName = BibleConfiguratorProgramClassName;
                        //loadInSameProcess = true;
                        break;
                    case "UnlockAllBible":
                        path = Path.Combine(Utils.GetCurrentDirectory(), BibleConfiguratorPath);
                        args = "-unlockAllBible";
                        programClassName = BibleConfiguratorProgramClassName;
                        //loadInSameProcess = true;
                        break;
                    case "SearchInDictionaries":
                        path = Path.Combine(Utils.GetCurrentDirectory(), BibleConfiguratorPath);
                        programClassName = BibleConfiguratorProgramClassName;
                        args = "-searchInDictionaries";
                        loadInSameProcess = true;
                        break;
                }

                RunProgram(path, programClassName, args, loadInSameProcess);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
		}

        private void RunProgram(string programPath, string programClassName, string args, bool loadInSameProcess = true)
        {
            if (loadInSameProcess)
            {
                try
                {
                    AssemblyLoader.InvokeMethod(new AssemblyLoader.MethodIdentifier()
                    {
                        AssemblyPath = programPath,
                        ClassName = programClassName,
                        MethodName = "RunFromAnotherApp"
                    }, args);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                    if (ex.InnerException != null)
                        MessageBox.Show(ex.InnerException.ToString());
                }
            }
            else
            {
                if (!string.IsNullOrEmpty(programPath))
                    Process.Start(programPath, args);
            }
        }

		/// <summary>
		/// Called from the loadImage="" parameter in ribbon.xml. Converts the images into IStreams
		/// </summary>
		/// <param name="imageName">The image="" parameter in ribbon.xml, i.e. the image name</param>
        public IStream GetImage(string imageName)
        {
            try
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
                    case "QuickAnalyze.png":
                        Properties.Resources.QuickAnalyze.Save(mem, ImageFormat.Png);
                        break;
                    case "UnlockFile.png":
                        Properties.Resources.UnlockFile.Save(mem, ImageFormat.Png);
                        break;
                    case "UnlockFolder.png":
                        Properties.Resources.UnlockFolder.Save(mem, ImageFormat.Png);
                        break;
                    case "Dictionary.png":
                        Properties.Resources.Dictionary.Save(mem, ImageFormat.Png);
                        break;
                }

                return new CCOMStreamWrapper(mem);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return null;
            }
        }

		#endregion
	}
}