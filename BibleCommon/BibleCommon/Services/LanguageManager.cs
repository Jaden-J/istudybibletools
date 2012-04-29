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
                    SettingsManager.Instance.Language != 0
                    ? SettingsManager.Instance.Language
                    : Thread.CurrentThread.CurrentUICulture.LCID);
            }
        }

        private static int[] _localsLCID =
            {
                1049,
                1033
            };

        public static string[] GetLanguagesNames()
        {
            List<string> names = new List<string>();

            foreach (var local in _localsLCID)
            {
                names.Add(BibleCommon.Resources.Constants.ResourceManager.GetString("_LANGUAGE_NAME", new CultureInfo(local)));
            }

            return names.ToArray();
        }

        public static Dictionary<int, string> GetDisplayedNames()
        {
            Dictionary<int, string> names = new Dictionary<int, string>();

            foreach (var local in _localsLCID)
            {
                names.Add(local, BibleCommon.Resources.Constants.ResourceManager.GetString("_LANGUAGE_NAME", new CultureInfo(local)));
            }

            return names;
        }
    }
}
