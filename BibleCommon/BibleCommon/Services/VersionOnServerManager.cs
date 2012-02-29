using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BibleCommon.Consts;
using System.Xml.Linq;
using System.Xml.XPath;

namespace BibleCommon.Services
{
    public static class VersionOnServerManager
    {
        public static bool NeedToUpdate()
        {
            bool result = false;

            if (SettingsManager.Instance.NewVersionOnServer != null
                   && SettingsManager.Instance.NewVersionOnServer > SettingsManager.Instance.CurrentVersion)
                result = true;
            else
            {
                Version newVersion = null;

                if (SettingsManager.Instance.NewVersionOnServerLatestCheckTime.HasValue)
                {
                    if (SettingsManager.Instance.NewVersionOnServerLatestCheckTime.Value.Add(Constants.NewVersionCheckPeriod) < DateTime.Now)
                        newVersion = TryToGetVersionOnServer();
                }
                else
                    newVersion = TryToGetVersionOnServer();

                if (newVersion != null)
                {
                    SettingsManager.Instance.NewVersionOnServerLatestCheckTime = DateTime.Now;

                    if (newVersion != SettingsManager.Instance.NewVersionOnServer)
                    {
                        SettingsManager.Instance.NewVersionOnServer = newVersion;

                        if (newVersion > SettingsManager.Instance.CurrentVersion)
                            result = true;
                    }

                    SettingsManager.Instance.Save();
                }
            }

            return result;
        }

        private static Version TryToGetVersionOnServer()
        {
            Version result = null;

            try
            {                
                XDocument xDoc = XDocument.Load(Constants.NewVersionOnServerFileUrl);

                XElement latestVersion = xDoc.Root.XPathSelectElement("LatestVersion_IStudyBibleTools_RU");
                if (latestVersion != null)
                {
                    result = new Version(latestVersion.Value);
                }                
            }
            catch (Exception)
            {
                //todo: log it in system log (ont in common log)
            }

            return result;
        }
    }
}
