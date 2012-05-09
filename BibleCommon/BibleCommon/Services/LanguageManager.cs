using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Globalization;
using System.Threading;
using System.Windows.Forms;

namespace BibleCommon.Services
{
    public static class LanguageManager
    {
        public static CultureInfo UserLanguage
        {
            get
            {
                int lcid = DefaultLCID;

                if (SettingsManager.Instance.Language != 0)
                    lcid = SettingsManager.Instance.Language;
                else if (_localesLCID.Contains(Thread.CurrentThread.CurrentUICulture.LCID))
                    lcid = Thread.CurrentThread.CurrentUICulture.LCID;                

                return new CultureInfo(lcid);
            }
        }


        public static readonly int DefaultLCID = 1033;

        private static int[] _localesLCID =
            {
                1049,
                1033
            };

        public static string[] GetLanguagesNames()
        {
            List<string> names = new List<string>();

            foreach (var local in _localesLCID)
            {
                names.Add(BibleCommon.Resources.Constants.ResourceManager.GetString("_LANGUAGE_NAME", new CultureInfo(local)));
            }

            return names.ToArray();
        }

        public static Dictionary<int, string> GetDisplayedNames()
        {
            Dictionary<int, string> names = new Dictionary<int, string>();

            foreach (var local in _localesLCID)
            {
                names.Add(local, BibleCommon.Resources.Constants.ResourceManager.GetString("_LANGUAGE_NAME", new CultureInfo(local)));
            }

            return names;
        }

        public static void SetFormUICulture(this Form form)
        {
            // настройки не влияют на работу в дизайнере
            if (form.Site == null || !form.Site.DesignMode)
            {
                // устанавливаем культуру обязательно до InitializeComponent();
                Thread.CurrentThread.CurrentUICulture = LanguageManager.UserLanguage;
            }
        }
    }
}
