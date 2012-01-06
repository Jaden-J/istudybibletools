using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace BibleVerseLinker
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                VerseLinkManager vlManager = new VerseLinkManager();

                if (args.Length == 1)
                    vlManager.DescriptionPageName = args[0];

                vlManager.Do();
            }
            catch (Exception ex)
            {
                Logger.LogError(ex.Message);
            }

            if(Logger.WasLogged)
               Console.ReadKey();
        }
    }
}
