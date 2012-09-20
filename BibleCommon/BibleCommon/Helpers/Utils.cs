using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.IO;
using BibleCommon.Consts;
using System.Xml.Serialization;


namespace BibleCommon.Helpers
{
    public static class Utils
    {
        public static string GetCurrentDirectory()
        {
            var assembly = Assembly.GetExecutingAssembly().CodeBase;
            var uri = new Uri(assembly);
            return Path.GetDirectoryName(uri.LocalPath);
        }

        public static int? GetVerseNumber(string textElementValue)
        {
            int? result = null;
            if (textElementValue.StartsWith("<a href"))
            {
                string searchPattern = ">";
                int i = textElementValue.IndexOf(searchPattern);
                if (i != -1)
                    result = StringUtils.GetStringFirstNumber(textElementValue, i + searchPattern.Length);
            }
            else
                result = StringUtils.GetStringFirstNumber(textElementValue);

            return result;
        }

        public static string GetTempFolderPath()
        {
            string s = Path.Combine(GetProgramDirectory(), Constants.TempDirectory);
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
    }
}
