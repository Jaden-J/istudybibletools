using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.IO;
using System.Windows.Forms;

namespace ISBTCommandHandler
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main(params string[] args)
        {
            if (args.Length == 0)
            {
                ServiceBase[] ServicesToRun;
                ServicesToRun = new ServiceBase[] 
			    { 
				    new LinkHandler() 
			    };
                ServiceBase.Run(ServicesToRun);
            }
            else
            {
                TryToSendMessage();       
            }
        }

        private static void TryToSendMessage()
        {
            try
            {
                if (AppMessenger.CheckPrevInstance())
                {
                    
                }
            }
            catch (Exception ex)
            {
                File.WriteAllText("c:\\log.txt", ex.ToString());
            }
        }

        //protected override void WndProc(ref Message m)
        //{
        //    if (m.Msg == AppMessenger.WM_COPYDATA)
        //    {
        //        string command =
        //        AppMessenger.ProcessWM_COPYDATA(m);
        //        if (command != null)
        //        {

        //            processCommandLine(command);
        //            return;
        //        }
        //    }
        //    base.WndProc(ref m);
        //}
    }
}
