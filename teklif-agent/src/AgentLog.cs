using System;
using System.IO;

namespace TeklifAgent
{
    internal static class AgentLog
    {
        private static readonly object Lock = new object();

        public static void Info(string msg)
        {
            Write("INFO", msg);
        }

        public static void Error(string msg)
        {
            Write("ERR ", msg);
        }

        private static void Write(string level, string msg)
        {
            try
            {
                Directory.CreateDirectory(AgentConfig.ConfigDir);
                var line = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " [" + level + "] " + msg;
                lock (Lock)
                {
                    File.AppendAllText(
                        Path.Combine(AgentConfig.ConfigDir, "agent.log"),
                        line + Environment.NewLine);
                }
            }
            catch { }
        }
    }
}
