using System;
using System.IO;
using System.Net;
using System.Text;
using System.Web.Script.Serialization;

namespace TeklifAgent
{
    public sealed class ApiClient
    {
        private readonly AgentConfig _cfg;
        private readonly JavaScriptSerializer _json = new JavaScriptSerializer();

        static ApiClient()
        {
            try
            {
                ServicePointManager.SecurityProtocol =
                    (SecurityProtocolType)3072; // Tls12
            }
            catch { }
        }

        public ApiClient(AgentConfig cfg)
        {
            _cfg = cfg;
        }

        public void SendHeartbeat()
        {
            var body = new
            {
                mac = _cfg.Mac,
                hostname = _cfg.Hostname,
                user = _cfg.User,
                excelVersion = _cfg.ExcelVersion,
                timestamp = DateTime.UtcNow.ToString("o")
            };
            PostJson(_cfg.NormalizedApiUrl() + "heartbeat/", _json.Serialize(body));
            AgentLog.Info("heartbeat OK mac=" + (_cfg.Mac ?? "?"));
        }

        public PendingCommand ClaimPendingCommand()
        {
            var macEnc = Uri.EscapeDataString(NormalizeMac(_cfg.Mac ?? ""));
            var url = _cfg.NormalizedApiUrl() + "commands/pending/" + macEnc + "/";
            AgentLog.Info("poll " + url);
            var resp = GetJson(url);
            if (string.IsNullOrEmpty(resp)) return null;

            AgentLog.Info("poll resp: " + (resp.Length > 200 ? resp.Substring(0, 200) : resp));

            var root = _json.Deserialize<ApiRoot<PendingCommandDto>>(resp);
            if (root == null || !root.success || root.data == null) return null;

            var d = root.data;
            if (d.id <= 0 || string.IsNullOrEmpty(d.moduleName)) return null;

            return new PendingCommand
            {
                Id = d.id,
                ModuleName = d.moduleName,
                Param = d.param
            };
        }

        private static string NormalizeMac(string mac)
        {
            return mac.Trim().ToUpperInvariant().Replace("-", ":");
        }

        public void MarkCommandDone(int id, string result, string errorMsg)
        {
            var status = string.IsNullOrEmpty(errorMsg) ? "done" : "error";
            var body = new
            {
                status,
                result = result ?? "",
                errorMsg = errorMsg ?? ""
            };
            PatchJson(_cfg.NormalizedApiUrl() + "commands/" + id + "/", _json.Serialize(body));
        }

        private string GetJson(string url)
        {
            var req = (HttpWebRequest)WebRequest.Create(url);
            req.Method = "GET";
            req.Timeout = 15000;
            req.ReadWriteTimeout = 15000;
            req.UserAgent = "TeklifAgent/1.0";
            using (var res = (HttpWebResponse)req.GetResponse())
            using (var sr = new StreamReader(res.GetResponseStream(), Encoding.UTF8))
                return sr.ReadToEnd();
        }

        private void PostJson(string url, string json)
        {
            SendJson(url, "POST", json);
        }

        private void PatchJson(string url, string json)
        {
            SendJson(url, "PATCH", json);
        }

        private void SendJson(string url, string method, string json)
        {
            var req = (HttpWebRequest)WebRequest.Create(url);
            req.Method = method;
            req.ContentType = "application/json; charset=utf-8";
            req.Timeout = 20000;
            req.ReadWriteTimeout = 20000;
            req.UserAgent = "TeklifAgent/1.0";
            var bytes = Encoding.UTF8.GetBytes(json);
            req.ContentLength = bytes.Length;
            using (var s = req.GetRequestStream())
                s.Write(bytes, 0, bytes.Length);
            using (var res = (HttpWebResponse)req.GetResponse())
            using (var sr = new StreamReader(res.GetResponseStream(), Encoding.UTF8))
                sr.ReadToEnd();
        }

        private class ApiRoot<T>
        {
            public bool success { get; set; }
            public T data { get; set; }
        }

        private class PendingCommandDto
        {
            public int id { get; set; }
            public string moduleName { get; set; }
            public string param { get; set; }
        }
    }

    public sealed class PendingCommand
    {
        public int Id { get; set; }
        public string ModuleName { get; set; }
        public string Param { get; set; }
    }
}
