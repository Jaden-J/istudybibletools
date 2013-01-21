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


namespace BibleCommon.Helpers
{
    public static class Utils
    {
        public static Version GetProgramVersion()
        {
            var assembly = Assembly.GetExecutingAssembly();
            return assembly.GetName().Version;
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

        public static string GetProgramDirectory()
        {
            string directoryPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), Consts.Constants.ToolsName);

            if (!Directory.Exists(directoryPath))
                Directory.CreateDirectory(directoryPath);

            return directoryPath;
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
            XmlSerializer serializer = new XmlSerializer(data.GetType());
            using (var fs = new FileStream(filePath, FileMode.Create))
            {
                serializer.Serialize(fs, data);
                fs.Flush();
            }
        }

        public static T LoadFromXmlFile<T>(string filePath)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(T));
            return (T)serializer.Deserialize(new MemoryStream(File.ReadAllBytes(filePath)));
        }

        public static T LoadFromXmlString<T>(string value)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(T));
            using (MemoryStream ms = new MemoryStream())
            {
                using (StreamWriter sw = new StreamWriter(ms))
                {
                    sw.WriteLine(value);
                    sw.Flush();
                    ms.Position = 0;

                    return (T)serializer.Deserialize(ms);
                }                
            }           
        }

        public static Encoding GetFileEncoding(string filePath)
        {
            System.Text.Encoding result = null;
            using (FileStream fs = new System.IO.FileStream(filePath,
                FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                if (fs.CanSeek)
                {
                    byte[] bom = new byte[4]; // Get the byte-order mark, if there is one 
                    fs.Read(bom, 0, 4);
                    if ((bom[0] == 0xef && bom[1] == 0xbb && bom[2] == 0xbf)
                        || (bom[0] == 47 && bom[1] == 47 && bom[2] == 32 && bom[3] == 208)
                        || (bom[0] == 60 && bom[1] == 109 && bom[2] == 101 && bom[3] == 116)
                        || (bom[0] == 60 && bom[1] == 116 && bom[2] == 105 && bom[3] == 116))  // utf-8 
                    {
                        result = System.Text.Encoding.UTF8;
                    }
                    else if ((bom[0] == 0xff && bom[1] == 0xfe)   // ucs-2le, ucs-4le, and ucs-16le 
                        || (bom[0] == 0xfe && bom[1] == 0xff) // utf-16 and ucs-2 
                        || (bom[0] == 0 && bom[1] == 0 && bom[2] == 0xfe && bom[3] == 0xff)) // ucs-4 
                    {
                        result = System.Text.Encoding.Unicode;
                    }
                    else
                    {
                        result = System.Text.Encoding.Default;
                    }

                    // Now reposition the file cursor back to the start of the file 
                    fs.Seek(0, System.IO.SeekOrigin.Begin);
                }
                else
                {
                    // The file cannot be randomly accessed, so you need to decide what to set the default to 
                    // based on the data provided. If you're expecting data from a lot of older applications, 
                    // default your encoding to Encoding.ASCII. If you're expecting data from a lot of newer 
                    // applications, default your encoding to Encoding.Unicode. Also, since binary files are 
                    // single byte-based, so you will want to use Encoding.ASCII, even though you'll probably 
                    // never need to use the encoding then since the Encoding classes are really meant to get 
                    // strings from the byte array that is the file. 

                    result = System.Text.Encoding.Default;
                }
            }

            return result;
        }

        public static string GetHexError(Error error)
        {
            return string.Format("0x{0}", Convert.ToString((int)error, 16));            
        }

        public static void Wait(Func<bool> checkIfExternalProcessAborted)
        {
            for (var i = 0; i < 3; i++)
            {
                Thread.Sleep(1000);
                if (checkIfExternalProcessAborted != null)
                {
                    if (checkIfExternalProcessAborted())
                        throw new ProcessAbortedByUserException();
                }
                System.Windows.Forms.Application.DoEvents();
            }
        }
    }
}
