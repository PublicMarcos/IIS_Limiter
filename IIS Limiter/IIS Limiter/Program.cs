using Microsoft.Web.Administration;
using System;
using System.Collections.Generic;
using System.IO;
using System.ServiceProcess;
using System.Threading;

namespace IIS_Limiter
{
    class Program
    {
        private static void Main(string[] args)
        {
            string ConfigPath = string.Empty, LogPath = string.Empty;
            var CheckInterval = TimeSpan.Parse("01:00:00");
            for (int i = 0; i < args.Length; i++)
            {
                if (args[i].StartsWith("-config=", StringComparison.InvariantCultureIgnoreCase))
                {
                    ConfigPath = args[i].Split(new string[1] { "-config=" }, StringSplitOptions.None)[1];
                }
                else if (args[i].StartsWith("-log=", StringComparison.InvariantCultureIgnoreCase))
                {
                    LogPath = args[i].Split(new string[1] { "-log=" }, StringSplitOptions.None)[1];
                }
                else if (args[i].StartsWith("-checkinterval=", StringComparison.InvariantCultureIgnoreCase))
                {
                    TimeSpan.TryParse(args[i].Split(new string[1] { "-checkinterval=" }, StringSplitOptions.None)[1], out CheckInterval);
                }
            }
            if (string.IsNullOrWhiteSpace(ConfigPath) || string.IsNullOrWhiteSpace(LogPath))
            {
                Console.WriteLine("Falsche Startparameter");
                Environment.Exit(87);
            }
            else
            {
                MainConfig = new Config(ConfigPath, LogPath);
            }

            while (true)
            {
                for (int i = 0; i < MainConfig.Items.Count; i++)
                {
                    var CurSpeed = GetAdapterSpeed(MainConfig.Items[i].Adaptername);
                    if (MainConfig.Items[i].LastAdapterSpeed != CurSpeed)
                    {
                        if ((CurSpeed >= MainConfig.Items[i].SpeedTriggerFrom) && (CurSpeed <= MainConfig.Items[i].SpeedTriggerTo))
                        {
                            if (LimitIISSpeed(MainConfig.Items[i].WebSiteName, MainConfig.Items[i].MaxBytesPerSec, MainConfig.Items[i].MaxConn, MainConfig.Items[i].Timeout, MainConfig.Items[i].IISServiceName))
                            {
                                ErrorReport(string.Format("Datum: {0:G}", DateTime.Now) + " - Änderungen wurden vorgenommen");
                            }
                            else
                            {
                                ErrorReport(string.Format("Datum: {0:G}", DateTime.Now) + " - Unbekannter Fehler");
                            }
                        }
                    }
                    MainConfig.Items[i].LastAdapterSpeed = CurSpeed;
                }
                Thread.Sleep((int)CheckInterval.TotalMilliseconds);
            }
        }

        private static Config MainConfig { get; set; }

        private static void ErrorReport(string Message)
        {
            Console.WriteLine(Message);
            File.AppendAllLines(MainConfig.LogPath, new string[1] { Message });
        }

        private static int GetAdapterSpeed(string Adaptername)
        {
            // The enum value of `AF_INET` will select only IPv4 adapters.
            // You can change this to `AF_INET6` for IPv6 likewise
            // And `AF_UNSPEC` for either one
            foreach (IPIntertop.IP_ADAPTER_ADDRESSES net in IPIntertop.GetIPAdapters(IPIntertop.FAMILY.AF_UNSPEC))
            {
                if (net.FriendlyName.ToLowerInvariant().Contains(Adaptername.ToLowerInvariant()))
                {
                    return (int)(net.TrasmitLinkSpeed / 1000000);
                }
            }
            ErrorReport("Adapter '" + Adaptername + "' nicht gefunden");
            return 0;
        }

        private static bool LimitIISSpeed(string WebSiteName, int MaxBytesPerSec, int MaxConn, TimeSpan Timeout, string IISServiceName)
        {
            try
            {
                using (var serverManager = new ServerManager())
                {
                    var config = serverManager.GetApplicationHostConfiguration();
                    var sitesSection = config.GetSection("system.applicationHost/sites");
                    var sitesCollection = sitesSection.GetCollection();

                    var siteElement = FindElement(sitesCollection, "site", "name", WebSiteName);
                    if (ReferenceEquals(siteElement, null))
                    {
                        ErrorReport("Seite '" + WebSiteName + "' nicht gefunden");
                        return false;
                    }

                    var limitsElement = siteElement.GetChildElement("limits");
                    limitsElement["maxBandwidth"] = MaxBytesPerSec;
                    limitsElement["maxConnections"] = MaxConn;
                    limitsElement["connectionTimeout"] = Timeout;

                    serverManager.CommitChanges();
                    var service = new ServiceController(IISServiceName);
                    service.Stop();
                    service.WaitForStatus(ServiceControllerStatus.Stopped, TimeSpan.Parse("00:01:00"));
                    service.Start();
                    service.WaitForStatus(ServiceControllerStatus.Running, TimeSpan.Parse("00:01:00"));
                    return true;
                }
            }
            catch
            {
                return false;
            }
        }

        private static ConfigurationElement FindElement(ConfigurationElementCollection collection, string elementTagName, params string[] keyValues)
        {
            foreach (ConfigurationElement element in collection)
            {
                if (element.ElementTagName.Equals(elementTagName, StringComparison.OrdinalIgnoreCase))
                {
                    bool matches = true;
                    for (int i = 0; i < keyValues.Length; i += 2)
                    {
                        object CurAttributeValue = element.GetAttributeValue(keyValues[i]);
                        string value = null;
                        if (!ReferenceEquals(CurAttributeValue, null))
                        {
                            value = CurAttributeValue.ToString();
                        }
                        if (!value.Equals(keyValues[i + 1], StringComparison.OrdinalIgnoreCase))
                        {
                            matches = false;
                            break;
                        }
                    }
                    if (matches)
                    {
                        return element;
                    }
                }
            }
            return null;
        }
    }

    internal sealed class Config
    {
        internal Config(string ConfigPath, string LogPath)
        {
            this.LogPath = LogPath;
            Items = new List<ConfigItem>();
            var RawData = File.ReadAllLines(ConfigPath);
            string Adaptername, WebSiteName, IISServiceName;
            int SpeedTriggerFrom = 0, SpeedTriggerTo = 0, MaxBytesPerSec = 0, MaxConn = 0;
            TimeSpan Timeout;
            for (int i = 0; i < RawData.Length; i++)
            {
                var RAWItems = RawData[i].Split(',');
                if ((RAWItems.Length > 6) && (RAWItems[1].Contains("-")))
                {
                    Adaptername = RAWItems[0];
                    WebSiteName = RAWItems[2];
                    IISServiceName = RAWItems[6];
                    var RAWTriggerPoints = RAWItems[1].Split('-');
                    if (int.TryParse(RAWTriggerPoints[0], out SpeedTriggerFrom) && int.TryParse(RAWTriggerPoints[1], out SpeedTriggerTo) && int.TryParse(RAWItems[3], out MaxBytesPerSec) && int.TryParse(RAWItems[4], out MaxConn) && TimeSpan.TryParse(RAWItems[5], out Timeout))
                    {
                        Items.Add(new ConfigItem(Adaptername, SpeedTriggerFrom, SpeedTriggerTo, WebSiteName, MaxBytesPerSec, MaxConn, Timeout, IISServiceName));
                    }
                }
            }
        }
        internal List<ConfigItem> Items { get; set; }
        internal string LogPath { get; set; }
    }

    internal sealed class ConfigItem
    {
        internal ConfigItem(string Adaptername, int SpeedTriggerFrom, int SpeedTriggerTo, string WebSiteName, int MaxBytesPerSec, int MaxConn, TimeSpan Timeout, string IISServiceName)
        {
            this.Adaptername = Adaptername;
            this.SpeedTriggerFrom = SpeedTriggerFrom;
            this.SpeedTriggerTo = SpeedTriggerTo;
            this.WebSiteName = WebSiteName;
            this.MaxBytesPerSec = MaxBytesPerSec;
            this.MaxConn = MaxConn;
            this.Timeout = Timeout;
            this.IISServiceName = IISServiceName;
            this.LastAdapterSpeed = 0;
        }

        internal string Adaptername { get; set; }
        internal int SpeedTriggerFrom { get; set; }
        internal int SpeedTriggerTo { get; set; }
        internal string WebSiteName { get; set; }
        internal int MaxBytesPerSec { get; set; }
        internal int MaxConn { get; set; }
        internal TimeSpan Timeout { get; set; }
        internal string IISServiceName { get; set; }
        internal int LastAdapterSpeed { get; set; }
    }
}
