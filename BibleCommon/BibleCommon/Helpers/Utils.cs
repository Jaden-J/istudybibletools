using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.IO;
using BibleCommon.Consts;
using System.Xml.Serialization;
using Microsoft.Office.Interop.OneNote;
using BibleCommon.Common;
using System.Threading;
using System.Text.RegularExpressions;
using System.Globalization;
using BibleCommon.Services;


namespace BibleCommon.Helpers
{
    public static class Utils
    {
        public static Version GetProgramVersion()
        {
            var assembly = Assembly.GetExecutingAssembly();
            return assembly.GetName().Version;
        }

        public static string GetProgramDirectory()
        {
            string directoryPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), Consts.Constants.ToolsName);

            if (!Directory.Exists(directoryPath))
                Directory.CreateDirectory(directoryPath);

            return directoryPath;
        }

        public static string GetCurrentDirectory()
        {
            var assembly = Assembly.GetExecutingAssembly().CodeBase;
            var uri = new Uri(assembly);
            return Path.GetDirectoryName(uri.LocalPath);
        }

        public static string GetTempFolderPath()
        {
            string s = Path.Combine(GetProgramDirectory(), Constants.TempDirectory);
            if (!Directory.Exists(s))
                Directory.CreateDirectory(s);

            return s;
        }

        public static string GetCacheFolderPath()
        {
            string s = Path.Combine(GetProgramDirectory(), Constants.CacheDirectory);
            if (!Directory.Exists(s))
                Directory.CreateDirectory(s);

            return s;
        }

        public static string GetNotesPagesFolderPath()
        {
            string s = Path.Combine(GetProgramDirectory(), Constants.NotesPagesDirectory);
            if (!Directory.Exists(s))
                Directory.CreateDirectory(s);

            return s;
        }

        public static string GetAnalyzedVersesFolderPath()
        {
            string s = Path.Combine(SettingsManager.Instance.FolderPath_BibleNotesPages, Constants.AnalyzedVersesDirectory);
            if (!Directory.Exists(s))
                Directory.CreateDirectory(s);

            return s;
        }

        public static string GetNewDirectoryPath(string folderPath)
        {
            string result = folderPath;
            for (int i = 0; i < 200; i++)
            {
                result = folderPath + (i > 0 ? " (" + i.ToString() + ")" : string.Empty);

                if (!Directory.Exists(result))
                    return result;
            }

            return folderPath;
        }

        public static void SaveToXmlFile(object data, string filePath)
        {
            var serializer = XmlSerializerCache.GetXmlSerializer(data.GetType());
            using (var fs = new FileStream(filePath, FileMode.Create))
            {
                serializer.Serialize(fs, data);
                fs.Flush();
            }
        }

        public static T LoadFromXmlFile<T>(string filePath)
        {
            var serializer = XmlSerializerCache.GetXmlSerializer(typeof(T));
            return (T)serializer.Deserialize(new MemoryStream(File.ReadAllBytes(filePath)));
        }

        public static T LoadFromXmlString<T>(string value)
        {
            var serializer = XmlSerializerCache.GetXmlSerializer(typeof(T));
            using (var ms = new MemoryStream())
            {
                using (var sw = new StreamWriter(ms))
                {
                    sw.WriteLine(value);
                    sw.Flush();
                    ms.Position = 0;

                    return (T)serializer.Deserialize(ms);
                }                
            }           
        }

        public static void WaitFor(int seconds, Func<bool> checkIfExternalProcessAborted = null)
        {
            for (var i = 0; i < seconds * 10; i++)
            {
                Thread.Sleep(100);
                if (checkIfExternalProcessAborted != null)
                {
                    if (checkIfExternalProcessAborted())
                        throw new ProcessAbortedByUserException();
                }
                System.Windows.Forms.Application.DoEvents();
            }
        }  

        public static byte[] ReadStream(Stream input)
        {
            byte[] buffer = new byte[16 * 1024];
            using (MemoryStream ms = new MemoryStream())
            {
                int read;
                while ((read = input.Read(buffer, 0, buffer.Length)) > 0)
                {
                    ms.Write(buffer, 0, read);
                }
                return ms.ToArray();
            }
        }

        public static void DoWithExceptionHandling(bool silent, Action action)
        {
            DoWithExceptionHandling(string.Empty, silent, action);
        }

        public static void DoWithExceptionHandling(string errorDescription, bool silent, Action action)
        {
            try
            {
                if (action != null)
                    action();
            }
            catch (Exception ex)
            {
                if (silent)
                    Logger.LogError(errorDescription, ex);
                else
                    FormLogger.LogError(errorDescription, ex);
            }
        }

        public static string GetUpdateProgramWebSitePageUrl()
        {
            var result = BibleCommon.Resources.Constants.DownloadPageUrl;

            try
            {
                var filePath = SystemUtils.GetOneNoteProgramFilePath();
                var is64Bit = SystemUtils.UnmanagedDllIs64Bit(filePath);

                result = string.Format("{0}&v={1}&x64={2}", result, (int)SystemConstants.VersionOfOneNote, is64Bit);
            }
            catch (Exception ex)
            {
                Logger.LogError(ex);                
            }

            return result;
        }

        public static DateTime ParseDateTime(string s)
        {   
            try
            {
                return DateTime.Parse(s, CultureInfo.InvariantCulture);                
            }
            catch (FormatException)
            {
                try
                {
                    return DateTime.Parse(s);
                }
                catch (FormatException)
                {
                    return DateTime.Parse(s, LanguageManager.GetCurrentCultureInfo());
                }
            }
        }
    }
}
