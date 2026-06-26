using System;
using System.IO;
using System.Management;

namespace TeklifAgent
{
    public static class BootSession
    {
        public static string GetBootSessionId()
        {
            try
            {
                using (var searcher = new ManagementObjectSearcher(
                    "SELECT LastBootUpTime FROM Win32_OperatingSystem"))
                {
                    foreach (ManagementObject obj in searcher.Get())
                    {
                        var raw = obj["LastBootUpTime"] as string;
                        if (!string.IsNullOrEmpty(raw))
                            return raw.Replace(":", "").Replace(".", "");
                    }
                }
            }
            catch { }

            return DateTime.UtcNow.ToString("yyyyMMdd");
        }

        /// <summary>Yeni Windows oturumunda excel-session.ready sifirlanir.</summary>
        public static void EnsureBootMarker()
        {
            try
            {
                Directory.CreateDirectory(AgentConfig.ConfigDir);
                var bootId = GetBootSessionId();
                var marker = AgentConfig.BootMarkerPath;
                var prev = File.Exists(marker) ? File.ReadAllText(marker).Trim() : "";
                if (!string.Equals(prev, bootId, StringComparison.Ordinal))
                {
                    if (File.Exists(AgentConfig.ExcelSessionReadyPath))
                        File.Delete(AgentConfig.ExcelSessionReadyPath);
                    File.WriteAllText(marker, bootId);
                    AgentLog.Info("yeni boot oturumu boot=" + bootId);
                }
            }
            catch (Exception ex)
            {
                AgentLog.Error("boot-marker: " + ex.Message);
            }
        }
    }
}
