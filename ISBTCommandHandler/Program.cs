using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using System.IO;

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
                File.WriteAllText("c:\\log.txt", args[0]);                
                Thread.Sleep(5000);
            }
        }
    }
}
