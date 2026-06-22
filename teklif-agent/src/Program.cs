using System;
using System.IO;
using System.Net;
using System.Threading;

namespace TeklifAgent
{
    internal static class Program
    {
        private static void Main(string[] args)
        {
            try
            {
                ServicePointManager.SecurityProtocol =
                    (SecurityProtocolType)3072;
            }
            catch { }

            if (args.Length > 0 && string.Equals(args[0], "--worker", StringComparison.OrdinalIgnoreCase))
            {
                RunWorkerMode();
                return;
            }

            Console.WriteLine("TeklifAgent — Excel uzak modul arka plan ajani");
            Console.WriteLine("  TeklifAgent.exe --worker");
        }

        private static void RunWorkerMode()
        {
            var cfg = AgentConfig.Load();
            if (cfg == null)
            {
                AgentLog.Error("config.json bulunamadi");
                Environment.Exit(2);
                return;
            }

            bool created;
            var mtx = new Mutex(true, "Global\\TeklifAgent_" + (cfg.Mac ?? "default"), out created);
            if (!created)
            {
                AgentLog.Info("Baska agent instance calisiyor, cikiliyor");
                Environment.Exit(0);
                return;
            }

            AgentLog.Info("TeklifAgent.exe --worker basladi");

            var worker = new AgentWorker();
            worker.RunBlocking(cfg);

            mtx.ReleaseMutex();
        }
    }
}
