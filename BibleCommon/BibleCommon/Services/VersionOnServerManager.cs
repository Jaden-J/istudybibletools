using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BibleCommon.Consts;
using System.Xml.Linq;
using System.Xml.XPath;
using System.Net;
using System.Xml;
using System.IO;

namespace BibleCommon.Services
{
    public class VersionOnServerManager
    {
        public Version NewVersion { get; set; }

        /// <summary>
        /// Версия релиза. Если текущая версия меньше этой версии, тогда показываем окно с информацией о релизе.
        /// </summary>
        public Version ReleaseMinVersion { get; set; }
        public string ReleaseInfo { get; set; }

        public bool NeedToShowReleaseInfo()
        {
            return NeedToUpdate()
                    && !string.IsNullOrEmpty(ReleaseInfo)
                    && ReleaseMinVersion > SettingsManager.Instance.CurrentVersion;
        }

        public bool NeedToUpdate()
        {
            bool result = false;

            if (SettingsManager.Instance.NewVersionOnServer != null
                   && SettingsManager.Instance.NewVersionOnServer > SettingsManager.Instance.CurrentVersion)
            {
                NewVersion = SettingsManager.Instance.NewVersionOnServer;
                ReleaseInfo = SettingsManager.Instance.ReleaseInfo;
                ReleaseMinVersion = SettingsManager.Instance.ReleaseMinVersion;
                result = true;
            }
            else
            {
                if (!SettingsManager.Instance.NewVersionOnServerLatestCheckTime.HasValue
                    || SettingsManager.Instance.NewVersionOnServerLatestCheckTime.Value.Add(Constants.NewVersionCheckPeriod) < DateTime.Now)
                    TryToGetVersionOnServer();

                if (NewVersion != null)
                {
                    SettingsManager.Instance.NewVersionOnServerLatestCheckTime = DateTime.Now;

                    if (NewVersion != SettingsManager.Instance.NewVersionOnServer)
                    {
                        SettingsManager.Instance.NewVersionOnServer = NewVersion;
                        SettingsManager.Instance.ReleaseInfo = ReleaseInfo;
                        SettingsManager.Instance.ReleaseMinVersion = ReleaseMinVersion;

                        if (NewVersion > SettingsManager.Instance.CurrentVersion)
                            result = true;
                    }

                    SettingsManager.Instance.Save();
                }
            }

            return result;
        }

        private void TryToGetVersionOnServer()
        {
            try
            {
                LanguageManager.SetThreadUICulture();
                var xDoc = Load(BibleCommon.Resources.Constants.NewVersionOnServerFileUrl);

                var latestVersionEl = xDoc.Root.XPathSelectElement("LatestVersion");
                if (latestVersionEl != null)
                    NewVersion = new Version(latestVersionEl.Value);

                var releaseMinVersionEl = xDoc.Root.XPathSelectElement("ReleaseMinVersion");
                if (releaseMinVersionEl != null)
                    ReleaseMinVersion = new Version(releaseMinVersionEl.Value);

                var versionInfoEl = xDoc.Root.XPathSelectElement("ReleaseInfo");
                if (versionInfoEl != null)
                    ReleaseInfo = versionInfoEl.Value;
              
            }
            catch (Exception)
            {
                //todo: log it in system log (not in common log)
            }
        }


        private static Stream LoadHttpStream(string fullServiceUrl)
        {
            var httpRequest = GetWebRequest(fullServiceUrl);
            var httpResponse = (HttpWebResponse)httpRequest.GetResponse();
            return httpResponse.GetResponseStream();
        }

        private static XDocument Load(string fullServiceUrl)
        {            
            using (var stream = LoadHttpStream(fullServiceUrl))
            {
                using (StreamReader reader = new StreamReader(stream))
                {
                    var xDoc = XDocument.Load(reader);
                    return xDoc;
                }
            }
        }

        private static HttpWebRequest GetWebRequest(string url)
        {
            ServicePointManager.ServerCertificateValidationCallback += (sender, certificate, chain, sslPolicyErrors) => true;

            var httpRequest = (HttpWebRequest)WebRequest.Create(url);

            httpRequest.Proxy = GetProxy();
            httpRequest.Method = WebRequestMethods.Http.Get;

            return httpRequest;
        }
        
        private static WebProxy GetProxy()
        {
#pragma warning disable 0618
            var proxy = WebProxy.GetDefaultProxy();
#pragma warning restore 0618

            if (proxy.Address != null)
            {
                proxy.Credentials = CredentialCache.DefaultNetworkCredentials;
                WebRequest.DefaultWebProxy = new WebProxy(proxy.Address, proxy.BypassProxyOnLocal, proxy.BypassList, proxy.Credentials);
            }
            return proxy;
        }
    }
}
