using System;
using System.IO;
using System.Web.Script.Serialization;

namespace TeklifAgent
{
    public sealed class AgentConfig
    {
        public string ApiBaseUrl { get; set; }
        public string Mac { get; set; }
        public string Hostname { get; set; }
        public string User { get; set; }
        public string ExcelVersion { get; set; }
        public int IntervalMinutes { get; set; }
        public bool Stop { get; set; }

        public static string ConfigDir
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "TeklifAgent");
            }
        }

        public static string ConfigPath
        {
            get { return Path.Combine(ConfigDir, "config.json"); }
        }

        public static string InstallDir
        {
            get { return ConfigDir; }
        }

        public static string DllPath
        {
            get { return Path.Combine(ConfigDir, "TeklifAgent.Com.dll"); }
        }

        public static string StopFlagPath
        {
            get { return Path.Combine(ConfigDir, "stop.flag"); }
        }

        /// <summary>HeartbeatPing ilk Excel acilisinda olusturur; yoksa agent ping atmaz.</summary>
        public static string ExcelSessionReadyPath
        {
            get { return Path.Combine(ConfigDir, "excel-session.ready"); }
        }

        public static string BootMarkerPath
        {
            get { return Path.Combine(ConfigDir, "current-boot.id"); }
        }

        public static bool IsExcelSessionReady()
        {
            return File.Exists(ExcelSessionReadyPath);
        }

        public void Save()
        {
            Directory.CreateDirectory(ConfigDir);
            var ser = new JavaScriptSerializer();
            File.WriteAllText(ConfigPath, ser.Serialize(this), System.Text.Encoding.UTF8);
        }

        public static AgentConfig Load()
        {
            if (!File.Exists(ConfigPath)) return null;
            var ser = new JavaScriptSerializer();
            return ser.Deserialize<AgentConfig>(File.ReadAllText(ConfigPath, System.Text.Encoding.UTF8));
        }

        public string NormalizedApiUrl()
        {
            var u = (ApiBaseUrl ?? "").Trim();
            if (u.Length == 0) u = "https://nextjs-teklif-sunucu.vercel.app/api/";
            if (!u.EndsWith("/")) u += "/";
            return u;
        }
    }
}
