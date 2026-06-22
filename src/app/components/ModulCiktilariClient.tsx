"use client";

import { useState, useEffect, useCallback } from "react";
import { formatTR } from "@/lib/date-utils";
import type { ModuleOutput, ModuleOutputSummary } from "@/lib/db";

interface Props {
  initialOutputs: ModuleOutput[];
  summary: ModuleOutputSummary[];
  allMacs: string[];
  allModuleNames: string[];
  allHostnames: { mac: string; hostname: string }[];
}

function ago(iso: string) {
  const diff = Date.now() - new Date(iso).getTime();
  const m = Math.floor(diff / 60000);
  if (m < 1)  return "Az önce";
  if (m < 60) return `${m}dk önce`;
  const h = Math.floor(m / 60);
  if (h < 24) return `${h}s önce`;
  return `${Math.floor(h / 24)}g önce`;
}

function JsonViewer({ data }: { data: Record<string, unknown> }) {
  const [expanded, setExpanded] = useState(false);
  const entries = Object.entries(data);
  const preview = entries.slice(0, 3);
  const hasMore = entries.length > 3;

  return (
    <div style={{ fontSize: 12 }}>
      {(expanded ? entries : preview).map(([k, v]) => (
        <div key={k} style={{ display: "flex", gap: 6, marginBottom: 2, flexWrap: "wrap" }}>
          <span style={{ color: "var(--accent)", fontWeight: 600, minWidth: 100, flexShrink: 0 }}>{k}:</span>
          <span style={{ color: "var(--text-dim)", wordBreak: "break-all" }}>
            {typeof v === "object" ? JSON.stringify(v) : String(v ?? "")}
          </span>
        </div>
      ))}
      {hasMore && (
        <button
          onClick={() => setExpanded(!expanded)}
          style={{ background: "none", border: "none", color: "var(--accent)", cursor: "pointer", fontSize: 11, padding: "2px 0", marginTop: 2 }}
        >
          {expanded ? "▲ Daha az göster" : `▼ +${entries.length - 3} alan daha`}
        </button>
      )}
    </div>
  );
}

export default function ModulCiktilariClient({
  initialOutputs,
  summary,
  allMacs,
  allModuleNames,
  allHostnames,
}: Props) {
  const [outputs, setOutputs]         = useState<ModuleOutput[]>(initialOutputs);
  const [filterMac, setFilterMac]     = useState("");
  const [filterMod, setFilterMod]     = useState("");
  const [filterText, setFilterText]   = useState("");
  const [loading, setLoading]         = useState(false);
  const [activeTab, setActiveTab]     = useState<"list" | "summary">("list");
  const [deleting, setDeleting]       = useState<number | null>(null);

  const hostnameMap = Object.fromEntries(allHostnames.map((h) => [h.mac, h.hostname]));

  const fetchOutputs = useCallback(async () => {
    setLoading(true);
    const sp = new URLSearchParams();
    if (filterMac) sp.set("mac", filterMac);
    if (filterMod) sp.set("moduleName", filterMod);
    sp.set("limit", "300");
    const res = await fetch(`/api/module-output?${sp}`);
    const json = await res.json();
    if (json.success) setOutputs(json.data);
    setLoading(false);
  }, [filterMac, filterMod]);

  useEffect(() => {
    fetchOutputs();
  }, [fetchOutputs]);

  const filtered = outputs.filter((o) => {
    if (!filterText) return true;
    const q = filterText.toLowerCase();
    return (
      o.mac.toLowerCase().includes(q) ||
      o.moduleName.toLowerCase().includes(q) ||
      (o.hostname ?? "").toLowerCase().includes(q) ||
      (o.firmaAdi ?? "").toLowerCase().includes(q) ||
      JSON.stringify(o.output).toLowerCase().includes(q)
    );
  });

  const handleDelete = async (id: number) => {
    if (!confirm("Bu kaydı silmek istediğinize emin misiniz?")) return;
    setDeleting(id);
    await fetch(`/api/module-output/${id}`, { method: "DELETE" });
    setOutputs((prev) => prev.filter((o) => o.id !== id));
    setDeleting(null);
  };

  const categoryColors: Record<string, string> = {
    donanim: "#6366f1", registry: "#f59e0b", internet: "#3b82f6",
    powershell: "#10b981", bildirim: "#ec4899", uzman: "#8b5cf6",
    dosya: "#14b8a6", excel: "#22c55e", default: "#6b7280",
  };

  const getCatColor = (name: string) => {
    const lower = name.toLowerCase();
    for (const [key, color] of Object.entries(categoryColors)) {
      if (lower.includes(key)) return color;
    }
    return categoryColors.default;
  };

  return (
    <main style={{ padding: "24px", maxWidth: 1300, margin: "0 auto" }}>
      {/* Header */}
      <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 24 }}>
        <span style={{ fontSize: 28 }}>📤</span>
        <div>
          <h1 style={{ margin: 0, fontSize: 22, fontWeight: 700 }}>Modül Çıktıları</h1>
          <p style={{ margin: 0, fontSize: 13, color: "var(--text-dim)" }}>
            VBA modüllerinin sunucuya gönderdiği veriler
          </p>
        </div>
        <div style={{ marginLeft: "auto", display: "flex", gap: 8 }}>
          <button
            onClick={fetchOutputs}
            disabled={loading}
            className="btn-primary"
            style={{ fontSize: 13, padding: "7px 16px" }}
          >
            {loading ? "⏳" : "🔄"} Yenile
          </button>
        </div>
      </div>

      {/* Özet Kartlar */}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(160px,1fr))", gap: 12, marginBottom: 24 }}>
        <div className="stat-card">
          <div className="stat-value">{outputs.length}</div>
          <div className="stat-label">Toplam Kayıt</div>
        </div>
        <div className="stat-card">
          <div className="stat-value">{new Set(outputs.map((o) => o.moduleName)).size}</div>
          <div className="stat-label">Farklı Modül</div>
        </div>
        <div className="stat-card">
          <div className="stat-value">{new Set(outputs.map((o) => o.mac)).size}</div>
          <div className="stat-label">Farklı Cihaz</div>
        </div>
        <div className="stat-card">
          <div className="stat-value">
            {outputs.length > 0 ? ago(outputs[0].createdAt) : "-"}
          </div>
          <div className="stat-label">Son Çıktı</div>
        </div>
      </div>

      {/* Sekmeler */}
      <div style={{ display: "flex", gap: 0, marginBottom: 20, borderBottom: "1px solid var(--border)" }}>
        {(["list", "summary"] as const).map((tab) => (
          <button
            key={tab}
            onClick={() => setActiveTab(tab)}
            style={{
              background: "none", border: "none", cursor: "pointer",
              padding: "8px 20px", fontSize: 14, fontWeight: 600,
              color: activeTab === tab ? "var(--accent)" : "var(--text-dim)",
              borderBottom: activeTab === tab ? "2px solid var(--accent)" : "2px solid transparent",
              marginBottom: -1,
            }}
          >
            {tab === "list" ? "📋 Kayıt Listesi" : "📊 Modül Özeti"}
          </button>
        ))}
      </div>

      {/* ÖZET TABLOSU */}
      {activeTab === "summary" && (
        <div className="card">
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
            <thead>
              <tr style={{ borderBottom: "1px solid var(--border)" }}>
                {["Modül", "Çalışma Sayısı", "Son Çalışma", "Son Cihaz", "Son Hostname"].map((h) => (
                  <th key={h} style={{ textAlign: "left", padding: "8px 12px", color: "var(--text-dim)", fontWeight: 600 }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {summary.map((s) => (
                <tr
                  key={s.moduleName}
                  style={{ borderBottom: "1px solid var(--border)", cursor: "pointer" }}
                  onClick={() => { setFilterMod(s.moduleName); setActiveTab("list"); }}
                >
                  <td style={{ padding: "8px 12px" }}>
                    <span style={{
                      display: "inline-block", padding: "2px 8px", borderRadius: 6,
                      background: getCatColor(s.moduleName) + "20",
                      color: getCatColor(s.moduleName), fontWeight: 600, fontSize: 12,
                    }}>
                      {s.moduleName}
                    </span>
                  </td>
                  <td style={{ padding: "8px 12px", fontWeight: 700 }}>{s.count.toLocaleString()}</td>
                  <td style={{ padding: "8px 12px", color: "var(--text-dim)" }}>
                    {formatTR(s.lastRunAt)}
                  </td>
                  <td style={{ padding: "8px 12px", fontFamily: "monospace", fontSize: 11, color: "var(--text-dim)" }}>
                    {s.lastMac ?? "-"}
                  </td>
                  <td style={{ padding: "8px 12px", color: "var(--text-dim)" }}>
                    {s.lastHostname ?? "-"}
                  </td>
                </tr>
              ))}
              {summary.length === 0 && (
                <tr>
                  <td colSpan={5} style={{ padding: 24, textAlign: "center", color: "var(--text-dim)" }}>
                    Henüz modül çıktısı yok. VBA modüllerinden PostToServer çağırın.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      )}

      {/* KAYIT LİSTESİ */}
      {activeTab === "list" && (
        <>
          {/* Filtreler */}
          <div style={{ display: "flex", gap: 10, marginBottom: 16, flexWrap: "wrap" }}>
            <select
              value={filterMod}
              onChange={(e) => setFilterMod(e.target.value)}
              style={{ padding: "7px 12px", borderRadius: 8, border: "1px solid var(--border)", background: "var(--card)", color: "var(--text)", fontSize: 13, minWidth: 180 }}
            >
              <option value="">Tüm Modüller</option>
              {allModuleNames.map((m) => (
                <option key={m} value={m}>{m}</option>
              ))}
            </select>

            <select
              value={filterMac}
              onChange={(e) => setFilterMac(e.target.value)}
              style={{ padding: "7px 12px", borderRadius: 8, border: "1px solid var(--border)", background: "var(--card)", color: "var(--text)", fontSize: 13, minWidth: 160 }}
            >
              <option value="">Tüm Cihazlar</option>
              {allMacs.map((m) => (
                <option key={m} value={m}>
                  {hostnameMap[m] ? `${hostnameMap[m]} (${m})` : m}
                </option>
              ))}
            </select>

            <input
              type="text"
              placeholder="🔍 Ara..."
              value={filterText}
              onChange={(e) => setFilterText(e.target.value)}
              style={{ padding: "7px 12px", borderRadius: 8, border: "1px solid var(--border)", background: "var(--card)", color: "var(--text)", fontSize: 13, minWidth: 200 }}
            />

            {(filterMac || filterMod || filterText) && (
              <button
                onClick={() => { setFilterMac(""); setFilterMod(""); setFilterText(""); }}
                style={{ padding: "7px 12px", borderRadius: 8, border: "1px solid var(--border)", background: "var(--card)", color: "var(--text-dim)", cursor: "pointer", fontSize: 13 }}
              >
                ✕ Temizle
              </button>
            )}

            <span style={{ marginLeft: "auto", color: "var(--text-dim)", fontSize: 13, alignSelf: "center" }}>
              {filtered.length} kayıt
            </span>
          </div>

          {/* Kartlar */}
          {loading ? (
            <div style={{ textAlign: "center", padding: 40, color: "var(--text-dim)" }}>Yükleniyor…</div>
          ) : filtered.length === 0 ? (
            <div style={{ textAlign: "center", padding: 40, color: "var(--text-dim)" }}>
              Henüz modül çıktısı yok.
              <br /><br />
              <span style={{ fontSize: 13 }}>
                VBA'dan bir modülde <code>zInternet.PostModuleOutput</code> çağırın
                veya <code>PostToServer.bas</code> modülünü kullanın.
              </span>
            </div>
          ) : (
            <div style={{ display: "grid", gap: 10 }}>
              {filtered.map((output) => (
                <div
                  key={output.id}
                  className="card"
                  style={{ padding: "14px 16px" }}
                >
                  <div style={{ display: "flex", alignItems: "flex-start", gap: 12, flexWrap: "wrap" }}>
                    {/* Modül renk badge */}
                    <span style={{
                      display: "inline-flex", alignItems: "center", padding: "3px 10px",
                      borderRadius: 20, fontSize: 11, fontWeight: 700, flexShrink: 0,
                      background: getCatColor(output.moduleName) + "20",
                      color: getCatColor(output.moduleName),
                      border: `1px solid ${getCatColor(output.moduleName)}40`,
                    }}>
                      {output.moduleName}
                    </span>

                    {/* Cihaz bilgisi */}
                    <div style={{ flex: 1, minWidth: 200 }}>
                      <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 6 }}>
                        <span style={{ fontSize: 12, color: "var(--text-dim)" }}>
                          🖥 <strong>{output.hostname ?? hostnameMap[output.mac] ?? output.mac}</strong>
                        </span>
                        {output.firmaAdi && (
                          <span style={{ fontSize: 12, color: "var(--text-dim)" }}>
                            🏢 {output.firmaAdi}
                          </span>
                        )}
                        <span style={{ fontSize: 11, color: "var(--text-dim)", fontFamily: "monospace" }}>
                          {output.mac}
                        </span>
                      </div>
                      <JsonViewer data={output.output} />
                    </div>

                    {/* Sağ: zaman + sil */}
                    <div style={{ display: "flex", flexDirection: "column", alignItems: "flex-end", gap: 6, flexShrink: 0 }}>
                      <span style={{ fontSize: 11, color: "var(--text-dim)" }}>
                        {ago(output.createdAt)}
                      </span>
                      <span style={{ fontSize: 10, color: "var(--text-dim)" }}>
                        {formatTR(output.createdAt)}
                      </span>
                      <button
                        onClick={() => handleDelete(output.id)}
                        disabled={deleting === output.id}
                        style={{
                          background: "none", border: "1px solid #ef444440", color: "#ef4444",
                          borderRadius: 6, padding: "3px 8px", cursor: "pointer", fontSize: 11,
                        }}
                      >
                        {deleting === output.id ? "⏳" : "🗑 Sil"}
                      </button>
                    </div>
                  </div>
                </div>
              ))}
            </div>
          )}
        </>
      )}

      {/* VBA Kullanım Kılavuzu */}
      <details style={{ marginTop: 32 }}>
        <summary style={{ cursor: "pointer", fontWeight: 600, fontSize: 14, padding: "8px 0", color: "var(--accent)" }}>
          📖 VBA'dan Nasıl Veri Gönderilir?
        </summary>
        <div className="card" style={{ marginTop: 12, padding: 16 }}>
          <p style={{ fontSize: 13, marginTop: 0, color: "var(--text-dim)" }}>
            <strong>Yöntem 1:</strong> <code>PostToServer</code> modülünü zInternet.RunRemoteCode ile çekin, ardından:
          </p>
          <pre style={{ background: "var(--bg)", padding: 12, borderRadius: 8, fontSize: 12, overflowX: "auto", margin: 0 }}>{`' Herhangi bir modülün sonunda çağırın:
Dim output As String
output = "{""key1"":""value1"",""key2"":42}"
Call PostModuleOutput("ModulAdı", output)`}</pre>
          <p style={{ fontSize: 13, marginTop: 12, color: "var(--text-dim)" }}>
            <strong>Yöntem 2:</strong> Direkt HTTP POST:
          </p>
          <pre style={{ background: "var(--bg)", padding: 12, borderRadius: 8, fontSize: 12, overflowX: "auto", margin: 0 }}>{`Dim http As Object
Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
http.Open "POST", baseUrl & "module-output", False
http.setRequestHeader "Content-Type", "application/json"
http.send "{""mac"":""XX:XX"",""moduleName"":""ModulAdi"",""output"":{""result"":""ok""}}"`}</pre>
        </div>
      </details>
    </main>
  );
}
