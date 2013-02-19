using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BibleCommon.Contracts;

namespace BibleCommon.Handlers
{
    public class RebuildDictionaryFileCacheHandler : IProtocolHandler
    {
        public string ProtocolName
        {
            get { return "isbtrdc:"; }
        }

        /// <summary>
        /// Доступно только после вызова ExecuteCommand()
        /// </summary>
        public string ModuleShortName { get; set; }

        public string GetCommandUrl(string moduleName)
        {
            return string.Format("{0}{1}", ProtocolName, moduleName);
        }

        public bool IsProtocolCommand(string[] args)
        {
            return args.Length > 0 && args[0].StartsWith(ProtocolName, StringComparison.OrdinalIgnoreCase);
        }

        public void ExecuteCommand(string[] args)
        {
            if (args.Length == 0)
                throw new ArgumentNullException("args");

            ModuleShortName = Uri.UnescapeDataString(args[0]
                                .Split(new char[] { ':' })[1]);                                

            // всё необходимое действие выполняется в BibleConfigurator
        }
    }
}
