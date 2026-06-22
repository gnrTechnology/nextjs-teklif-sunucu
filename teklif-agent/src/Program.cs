using System;
using System.IO;
using System.Threading;

namespace TeklifAgent
{
    /// <summary>Konsol modu: TeklifAgent.exe --worker (COM olmadan arka plan)</summary>
    internal static class Program
    {
        private static void Main(string[] args)
        {
            if (args.Length > 0 && string.Equals(args[0], "--worker", StringComparison.OrdinalIgnoreCase))
            {
                RunWorkerMode();
                return;
            }

            Console.WriteLine("TeklifAgent — Excel uzak modül arka plan ajanı");
            Console.WriteLine("  TeklifAgent.exe --worker   Arka plan döngüsü (config.json gerekir)");
            Console.WriteLine("  COM ProgId: TeklifAgent.Agent");
        }

        private static void RunWorkerMode()
        {
            var cfg = AgentConfig.Load();
            if (cfg == null)
            {
                Environment.Exit(2);
                return;
            }

            bool created;
            using (var mtx = new Mutex(true, "Global\\TeklifAgent_" + (cfg.Mac ?? "default"), out created))
            {
                if (!created)
                {
                    // Baska process calisiyor — config guncellendi, cik
                    Environment.Exit(0);
                    return;
                }

                if (File.Exists(AgentConfig.StopFlagPath))
                    File.Delete(AgentConfig.StopFlagPath);

                var worker = new AgentWorker();
                worker.Start(cfg);

                while (worker.IsRunning)
                    Thread.Sleep(2000);
            }
        }
    }
}
