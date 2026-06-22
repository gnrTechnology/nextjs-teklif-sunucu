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
                // Zaten çalışan ajan varsa yalnızca config güncelle
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
                _thread = new Thread(Loop) { IsBackground = true, Name = "TeklifAgentWorker" };
                _thread.Start();
            }
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
            var intervalMs = Math.Max(1, _cfg.IntervalMinutes) * 60000;

            // İlk döngüde beklemeden hemen çalış
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
                    var api = new ApiClient(_cfg);

                    api.SendHeartbeat();
                    ProcessCommands(api);
                    _lastError = "";
                }
                catch (Exception ex)
                {
                    _lastError = ex.Message;
                }

                if (!_running) break;

                // 1 sn parçalarda bekle — Stop hızlı yanıt versin
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
        }

        private void ProcessCommands(ApiClient api)
        {
            PendingCommand cmd;
            while ((cmd = api.ClaimPendingCommand()) != null)
            {
                string result;
                try
                {
                    if (!ExcelRunner.IsExcelRunning())
                    {
                        api.MarkCommandDone(cmd.Id, null, "Excel çalışmıyor — komut atlandı");
                        continue;
                    }

                    result = ExcelRunner.RunRemoteModule(cmd.ModuleName, cmd.Param);
                    if (result != null && result.StartsWith("ERR:"))
                        api.MarkCommandDone(cmd.Id, null, result.Substring(4));
                    else
                        api.MarkCommandDone(cmd.Id, result ?? "OK", null);
                }
                catch (Exception ex)
                {
                    api.MarkCommandDone(cmd.Id, null, ex.Message);
                }
            }
        }
    }
}
