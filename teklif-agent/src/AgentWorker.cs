using System;
using System.IO;
using System.Threading;

namespace TeklifAgent
{
    public sealed class AgentWorker
    {
        private readonly object _lock = new object();
        private Thread _thread;
        private volatile bool _running;
        private AgentConfig _cfg;
        private string _lastError = "";

        public bool IsRunning
        {
            get { return _running; }
        }

        public string LastError
        {
            get { return _lastError; }
        }

        public void Start(AgentConfig cfg)
        {
            if (cfg == null) throw new ArgumentNullException("cfg");
            lock (_lock)
            {
                if (_running)
                {
                    cfg.Save();
                    _cfg = cfg;
                    return;
                }

                if (File.Exists(AgentConfig.StopFlagPath))
                    File.Delete(AgentConfig.StopFlagPath);

                cfg.Stop = false;
                cfg.Save();
                _cfg = cfg;
                _running = true;
                _thread = new Thread(Loop)
                {
                    IsBackground = false,
                    Name = "TeklifAgentWorker"
                };
                _thread.Start();
            }
        }

        /// <summary>Exe modu — ana thread uzerinde calisir (process hayatta kalir)</summary>
        public void RunBlocking(AgentConfig cfg)
        {
            if (cfg == null) throw new ArgumentNullException("cfg");
            if (File.Exists(AgentConfig.StopFlagPath))
                File.Delete(AgentConfig.StopFlagPath);
            cfg.Stop = false;
            cfg.Save();
            _cfg = cfg;
            _running = true;
            AgentLog.Info("RunBlocking basladi mac=" + (cfg.Mac ?? "?"));
            Loop();
        }

        public void Stop()
        {
            lock (_lock)
            {
                StopInternal(true);
            }
        }

        public void SendPingNow()
        {
            var cfg = _cfg ?? AgentConfig.Load();
            if (cfg == null) throw new InvalidOperationException("Agent yapılandırması yok.");
            var api = new ApiClient(cfg);
            api.SendHeartbeat();
        }

        private void StopInternal(bool wait)
        {
            _running = false;
            try
            {
                Directory.CreateDirectory(AgentConfig.ConfigDir);
                File.WriteAllText(AgentConfig.StopFlagPath, "1");
            }
            catch { }

            if (_thread != null && wait)
            {
                if (!_thread.Join(5000))
                    _thread.Abort();
                _thread = null;
            }
        }

        private void Loop()
        {
            AgentLog.Info("Loop basladi");

            while (_running)
            {
                try
                {
                    if (File.Exists(AgentConfig.StopFlagPath))
                    {
                        _running = false;
                        break;
                    }

                    _cfg = AgentConfig.Load() ?? _cfg;
                    var intervalMs = Math.Max(1, _cfg.IntervalMinutes) * 60000;
                    var api = new ApiClient(_cfg);

                    // Heartbeat her zaman once — komut takilsa bile ping gider
                    try
                    {
                        api.SendHeartbeat();
                        _lastError = "";
                    }
                    catch (Exception ex)
                    {
                        _lastError = ex.Message;
                        AgentLog.Error("heartbeat: " + ex.Message);
                    }

                    // Komut: dongu basina en fazla 1, timeout ile
                    try
                    {
                        ProcessOneCommand(api);
                    }
                    catch (Exception ex)
                    {
                        AgentLog.Error("komut: " + ex.Message);
                    }

                    if (!_running) break;

                    var waited = 0;
                    while (_running && waited < intervalMs)
                    {
                        if (File.Exists(AgentConfig.StopFlagPath))
                        {
                            _running = false;
                            break;
                        }
                        Thread.Sleep(1000);
                        waited += 1000;
                    }
                }
                catch (Exception ex)
                {
                    _lastError = ex.Message;
                    AgentLog.Error("loop: " + ex.Message);
                    Thread.Sleep(5000);
                }
            }

            AgentLog.Info("Loop bitti");
        }

        private void ProcessOneCommand(ApiClient api)
        {
            var cmd = api.ClaimPendingCommand();
            if (cmd == null) return;

            AgentLog.Info("komut alindi id=" + cmd.Id + " modul=" + cmd.ModuleName);

            if (!ExcelRunner.IsExcelRunning())
            {
                api.MarkCommandDone(cmd.Id, null, "Excel calismiyor");
                return;
            }

            string result = null;
            string errMsg = null;

            var t = new Thread(delegate()
            {
                try
                {
                    result = ExcelRunner.RunRemoteModule(cmd.ModuleName, cmd.Param);
                    if (result != null && result.StartsWith("ERR:"))
                        errMsg = result.Substring(4);
                }
                catch (Exception ex)
                {
                    errMsg = ex.Message;
                }
            });
            t.IsBackground = true;
            t.Start();

            if (!t.Join(90000))
            {
                try { t.Abort(); } catch { }
                errMsg = "Komut zaman asimi (90sn)";
            }

            if (string.IsNullOrEmpty(errMsg))
                api.MarkCommandDone(cmd.Id, result ?? "OK", null);
            else
                api.MarkCommandDone(cmd.Id, null, errMsg);
        }
    }
}
