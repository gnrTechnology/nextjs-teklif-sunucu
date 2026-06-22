using System;
using System.Runtime.InteropServices;
using System.Web.Script.Serialization;

[assembly: ComVisible(true)]
[assembly: Guid("A1B2C3D4-E5F6-7890-ABCD-EF1234567890")]

namespace TeklifAgent
{
    [ComVisible(true)]
    [Guid("B2C3D4E5-F6A7-8901-BCDE-F12345678901")]
    [ProgId("TeklifAgent.Agent")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public sealed class AgentControl
    {
        private static readonly AgentWorker Worker = new AgentWorker();
        private static readonly JavaScriptSerializer Json = new JavaScriptSerializer();

        public void Start(string configJson)
        {
            var cfg = Json.Deserialize<AgentConfig>(configJson);
            if (cfg == null) throw new ArgumentException("Geçersiz config JSON");
            if (string.IsNullOrEmpty(cfg.Mac)) throw new ArgumentException("mac zorunlu");
            cfg.Save();
            Worker.Start(cfg);
        }

        public void Stop()
        {
            Worker.Stop();
        }

        public void SendPingNow()
        {
            Worker.SendPingNow();
        }

        public string GetStatus()
        {
            if (Worker.IsRunning) return "running";
            return "stopped";
        }

        public string GetLastError()
        {
            return Worker.LastError ?? "";
        }

        public string GetInstallDir()
        {
            return AgentConfig.ConfigDir;
        }

        public bool IsExcelRunning()
        {
            return ExcelRunner.IsExcelRunning();
        }
    }
}
