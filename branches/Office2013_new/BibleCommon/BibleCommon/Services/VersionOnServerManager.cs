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
                LanguageManager.SetThreadUICulture();
                XDocument xDoc = Load(BibleCommon.Resources.Constants.NewVersionOnServerFileUrl);

                XElement latestVersion = xDoc.Root.XPathSelectElement("LatestVersion_IStudyBibleTools_RU");
                if (latestVersion != null)
                {
                    result = new Version(latestVersion.Value);
                }                
            }
            catch (Exception)
            {
                //todo: log it in system log (not in common log)
            }

            return result;
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
