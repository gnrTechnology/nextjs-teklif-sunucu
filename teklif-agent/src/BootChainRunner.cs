using System;
using System.IO;
using System.Runtime.InteropServices;

namespace TeklifAgent
{
    /// <summary>
    /// Windows oturum acilisinda firma auto-start modullerini Excel uzerinden dagitir.
    /// </summary>
    public static class BootChainRunner
    {
        private static string BootFlagPath
        {
            get { return Path.Combine(AgentConfig.ConfigDir, "boot-chain.done"); }
        }

        public static void RunIfNeeded(AgentConfig cfg)
        {
            try
            {
                var bootId = BootSession.GetBootSessionId();
                if (IsDoneForBoot(bootId))
                {
                    AgentLog.Info("boot-chain zaten tamam boot=" + bootId);
                    return;
                }

                AgentLog.Info("boot-chain basliyor boot=" + bootId);
                EnsureExcelAndPollHost(cfg);

                var result = ExcelRunner.RunRemoteModule("AutoStartOnExcelOpen", cfg.NormalizedApiUrl());
                if (result != null && result.StartsWith("ERR:"))
                {
                    AgentLog.Error("boot-chain: " + result);
                    return;
                }

                Directory.CreateDirectory(AgentConfig.ConfigDir);
                File.WriteAllText(BootFlagPath, bootId);
                AgentLog.Info("boot-chain tamam");
            }
            catch (Exception ex)
            {
                AgentLog.Error("boot-chain: " + ex.Message);
            }
        }

        private static bool IsDoneForBoot(string bootId)
        {
            if (!File.Exists(BootFlagPath)) return false;
            return string.Equals(File.ReadAllText(BootFlagPath).Trim(), bootId, StringComparison.Ordinal);
        }

        private static void EnsureExcelAndPollHost(AgentConfig cfg)
        {
            object excelObj = null;
            try
            {
                if (!ExcelRunner.IsExcelRunning())
                {
                    var excelType = Type.GetTypeFromProgID("Excel.Application");
                    if (excelType == null)
                        throw new InvalidOperationException("Excel.Application bulunamadi");

                    excelObj = Activator.CreateInstance(excelType);
                    dynamic xl = excelObj;
                    xl.Visible = false;
                    xl.DisplayAlerts = false;
                    xl.ScreenUpdating = false;

                    TryLoadTeklifAddin(xl);
                    xl.Workbooks.Add();
                }
                else
                {
                    excelObj = Marshal.GetActiveObject("Excel.Application");
                }

                dynamic app = excelObj;
                var pollPath = Path.Combine(AgentConfig.ConfigDir, "TeklifPollHost.xlsm");
                if (!File.Exists(pollPath)) return;

                bool found = false;
                foreach (dynamic wb in app.Workbooks)
                {
                    try
                    {
                        var name = (string)wb.Name;
                        if (name.IndexOf("TeklifPollHost", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            found = true;
                            break;
                        }
                    }
                    catch { }
                }

                if (!found)
                {
                    dynamic opened = app.Workbooks.Open(pollPath, UpdateLinks: 0, ReadOnly: false);
                    try { opened.Windows[1].Visible = false; } catch { }
                }
            }
            finally
            {
                if (excelObj != null)
                    Marshal.ReleaseComObject(excelObj);
            }
        }

        private static void TryLoadTeklifAddin(dynamic xl)
        {
            try
            {
                foreach (dynamic ai in xl.AddIns)
                {
                    try
                    {
                        var n = (string)ai.Name;
                        if (n.IndexOf("teklif", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            if (!ai.Installed) ai.Installed = true;
                            return;
                        }
                    }
                    catch { }
                }

                var addinDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "Microsoft", "AddIns");
                var candidate = Path.Combine(addinDir, "teklif.xlam");
                if (File.Exists(candidate))
                {
                    dynamic ai = xl.AddIns.Add(candidate, false);
                    ai.Installed = true;
                }
            }
            catch (Exception ex)
            {
                AgentLog.Error("addin yukleme: " + ex.Message);
            }
        }
    }
}
