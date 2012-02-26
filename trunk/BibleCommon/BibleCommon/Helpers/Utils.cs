using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.IO;

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
    }
}
