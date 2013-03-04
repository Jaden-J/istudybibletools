using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BibleCommon.Contracts;
using BibleCommon.Common;
using BibleCommon.Services;

namespace BibleCommon.Handlers
{
    public class OpenNotesPageHandler : IProtocolHandler
    {
        private const string _protocolName = "isbtNotesPage:";

        public string ProtocolName
        {
            get { return _protocolName; }
        }

        /// <summary>
        /// Свойство доступно только после выполнения метода ExecuteCommand()
        /// </summary>
        public VersePointer Verse { get; set; }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="vp"></param>
        /// <param name="moduleName">может быть null</param>
        /// <returns></returns>
        public string GetCommandUrl(VersePointer vp, string moduleName)
        {
            return GetCommandUrlStatic(vp, moduleName);
        }

        public static string GetCommandUrlStatic(VersePointer vp, string moduleName)
        {
            return string.Format("{0}{1}/{2} {3};{4}", _protocolName, moduleName, vp.Book.Index, vp.VerseNumber, vp.OriginalVerseName);
        }

        public bool IsProtocolCommand(params string[] args)
        {
            return args.Length > 0 && args[0].StartsWith(ProtocolName, StringComparison.OrdinalIgnoreCase);
        }

        public void ExecuteCommand(params string[] args)
        {                
            try
            {
                var parts = args[0].Split(new char[] { ';', '&' });
                if (parts.Length < 2)
                    throw new ArgumentException(string.Format("Ivalid versePointer args: {0}", args[0]));            

                var verseString = Uri.UnescapeDataString(parts[1]);

                Verse = new VersePointer(verseString);

                if (!Verse.IsValid)                    
                    throw new Exception(BibleCommon.Resources.Constants.BibleVersePointerCanNotParseString);
            }
            catch (InvalidModuleException imEx)
            {
                FormLogger.LogError(BibleCommon.Resources.Constants.Error_SystemIsNotConfigured + Environment.NewLine + imEx.Message);
            }
            catch (Exception ex)
            {
                FormLogger.LogError(ex);
            }
        }

        string IProtocolHandler.GetCommandUrl(string args)
        {
            return string.Format("{0}:{1}", ProtocolName, args);
        }      
    }
}
