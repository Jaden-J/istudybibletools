using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BibleCommon.Contracts;
using BibleCommon.Common;
using BibleCommon.Services;
using System.IO;

namespace BibleCommon.Handlers
{
    public class OpenNotesPageHandler : IProtocolHandler
    {
        private const string _protocolName = "isbtNotesPage:";
        private const string _rubbishPageName = "detailed";

        public string ProtocolName
        {
            get { return _protocolName; }
        }

        /// <summary>
        /// Свойство доступно только после выполнения метода ExecuteCommand()
        /// </summary>
        public VersePointer Verse { get; set; }
        public NotesPageType NotesPageType { get; set; }

        public string GetVerseFilePath()
        {
            return GetNotesPageFilePath(Verse, NotesPageType);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="vp"></param>
        /// <param name="moduleName">может быть null</param>
        /// <returns></returns>
        public string GetCommandUrl(VersePointer vp, string moduleName, NotesPageType notesPageType)
        {
            return GetCommandUrlStatic(vp, moduleName, notesPageType);
        }

        public static string GetCommandUrlStatic(VersePointer vp, string moduleName, NotesPageType notesPageType)
        {
            return string.Format("{0}{1}/{2} {3}{4};{5};{6}",
                _protocolName,
                moduleName,
                vp.Book.Index,
                vp.Chapter.Value,
                !vp.IsChapter ? ":" + vp.VerseNumber : string.Empty,
                vp.GetFriendlyFullVerseName(), notesPageType);
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

                if (parts.Length == 3)
                {
                    var notesPageTypeString = Uri.UnescapeDataString(parts[2]);
                    NotesPageType = (NotesPageType)Enum.Parse(typeof(NotesPageType), notesPageTypeString);
                }
                else
                    NotesPageType = Common.NotesPageType.Verse;

                // дальнейшая обработка осуществляется в CommandForm.ProcessCommandLine()
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

        /// <summary>
        /// Если используем файловую систему для хранения сводных заметок
        /// </summary>
        /// <param name="vp"></param>
        /// <param name="notesPageType"></param>
        /// <returns></returns>
        public static string GetNotesPageFilePath(VersePointer vp, NotesPageType notesPageType)
        {
            var path =
                    Path.Combine(
                            Path.Combine(SettingsManager.Instance.FolderPath_BibleNotesPages, SettingsManager.Instance.ModuleShortName),
                            Path.Combine(string.Format("{0:00}. {1}", vp.Book.Index, vp.Book.Name), vp.Chapter.Value.ToString())
                            );

            string fileName;

            if (notesPageType == NotesPageType.Detailed)
                fileName = _rubbishPageName;
            else if (notesPageType == NotesPageType.Chapter)
                fileName = "0";            
            else
                fileName = vp.VerseNumber.ToString(); 

            return Path.Combine(path, fileName + ".htm");
        }

        string IProtocolHandler.GetCommandUrl(string args)
        {
            return string.Format("{0}:{1}", ProtocolName, args);
        }      
    }
}
