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
        public string ReleaseInfo { get; set; }

        public bool NeedToUpdate()
        {
            bool result = false;

            if (SettingsManager.Instance.NewVersionOnServer != null
                   && SettingsManager.Instance.NewVersionOnServer > SettingsManager.Instance.CurrentVersion)
            {
                NewVersion = SettingsManager.Instance.NewVersionOnServer;
                ReleaseInfo = SettingsManager.Instance.NewVersionInfo;
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
                        SettingsManager.Instance.NewVersionInfo = ReleaseInfo;

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

                var latestVersion = xDoc.Root.XPathSelectElement("LatestVersion");
                if (latestVersion != null)
                    NewVersion = new Version(latestVersion.Value);                

                var versionInfo = xDoc.Root.XPathSelectElement("LatestVersionInfo");
                if (versionInfo != null)
                    ReleaseInfo = versionInfo.Value;
              
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
