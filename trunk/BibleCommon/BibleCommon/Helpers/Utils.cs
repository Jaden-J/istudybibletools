﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.IO;
using BibleCommon.Consts;


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
    }
}