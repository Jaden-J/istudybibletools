using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Globalization;
using System.Threading;

namespace BibleCommon.Services
{
    public static class LanguageManager
    {
        public static CultureInfo UserLanguage
        {
            get
            {
                return new CultureInfo(
                    !string.IsNullOrEmpty(SettingsManager.Instance.Language)
                    ? SettingsManager.Instance.Language
                    : Thread.CurrentThread.CurrentUICulture.Name);
            }
        }

        private static string[] _locals =
            {
                "ru-RU",
                "en-US"
            };

        public static string[] GetLanguagesNames()
        {
            List<string> names = new List<string>();

            foreach (string local in _locals)
            {
                names.Add(BibleCommon.Resources.Constants.ResourceManager.GetString("_LANGUAGE_NAME", new CultureInfo(local)));
            }

            return names.ToArray();
        }

        public static Dictionary<string, string> GetDisplayedNames()
        {
            Dictionary<string, string> names = new Dictionary<string, string>();

            foreach (string local in _locals)
            {
                names.Add(local, BibleCommon.Resources.Constants.ResourceManager.GetString("_LANGUAGE_NAME", new CultureInfo(local)));
            }

            return names;
        }
    }
}
