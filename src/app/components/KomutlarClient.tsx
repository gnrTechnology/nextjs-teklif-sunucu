"use client";

import { useState, useEffect, useCallback } from "react";
import Link from "next/link";
import { formatTR, formatDurationSec, elapsedSecSince, timeAgo } from "@/lib/date-utils";
import type { ClientCommand } from "@/lib/db";
import type { FolderWatchHealth } from "@/lib/types";

const STATUS_CFG: Record<string, { label: string; color: string; bg: string }> = {
  pending: { label: "Bekliyor",  color: "#f59e0b", bg: "#f59e0b20" },
  running: { label: "Çalışıyor", color: "#3b82f6", bg: "#3b82f620" },
  done:    { label: "Tamamlandı",color: "#10b981", bg: "#10b98120" },
  error:   { label: "Hata",      color: "#ef4444", bg: "#ef444420" },
};

function StatusBadge({ status }: { status: string }) {
  const c = STATUS_CFG[status] ?? STATUS_CFG.pending;
  return (
    <span style={{
      fontSize: 11, padding: "2px 9px", borderRadius: 10, fontWeight: 600,
      background: c.bg, color: c.color,
    }}>{c.label}</span>
  );
}

export default function KomutlarClient({
  initial,
  allModuleNames,
  allMacs,
}: {
  initial: ClientCommand[];
  allModuleNames: string[];
  allMacs: string[];
}) {
  const [commands, setCommands] = useState<ClientCommand[]>(initial);
  const [filterStatus, setFilterStatus] = useState<string>("tümü");
  const [filterMac, setFilterMac]       = useState<string>("");
  const [expanded, setExpanded]         = useState<Set<number>>(new Set());

  /* Yeni komut formu */
  const [formMac, setFormMac]           = useState(allMacs[0] ?? "");
  const [formModule, setFormModule]     = useState("");
  const [formParam, setFormParam]       = useState("");
  const [formCustomMac, setFormCustomMac] = useState("");
  const [sending, setSending]           = useState(false);
  const [watchHealth, setWatchHealth]   = useState<Record<string, FolderWatchHealth>>({});
  const [, setTick]                     = useState(0);

  const refresh = useCallback(() => {
    fetch("/api/commands/?limit=200")
      .then((r) => r.json())
      .then((j) => { if (j.success) setCommands(j.data); });
  }, []);

  const refreshWatchHealth = useCallback(async (macs: string[]) => {
    const unique = [...new Set(macs.filter(Boolean))];
    if (unique.length === 0) return;
    const entries = await Promise.all(
      unique.map(async (m) => {
        const r = await fetch(`/api/folder-watch/?health=1&mac=${encodeURIComponent(m)}`);
        const j = await r.json();
        const h = j.success && j.data?.[0] ? (j.data[0] as FolderWatchHealth) : null;
        return [m, h] as const;
      }),
    );
    setWatchHealth((prev) => {
      const next = { ...prev };
      for (const [m, h] of entries) {
        if (h) next[m] = h;
      }
      return next;
    });
  }, []);

  /* 10sn'de bir yenile (pending/running görünür hale gelsin) */
  useEffect(() => {
    const id = setInterval(refresh, 10000);
    return () => clearInterval(id);
  }, [refresh]);

  /* running süresi canlı güncellensin */
  useEffect(() => {
    const hasRunning = commands.some((c) => c.status === "running");
    if (!hasRunning) return;
    const id = setInterval(() => setTick((t) => t + 1), 1000);
    return () => clearInterval(id);
  }, [commands]);

  /* WatchFolderServer running iken klasör sağlık durumu */
  useEffect(() => {
    const macs = commands
      .filter((c) => c.status === "running" && c.moduleName === "WatchFolderServer")
      .map((c) => c.mac);
    if (macs.length === 0) return;
    refreshWatchHealth(macs);
    const id = setInterval(() => refreshWatchHealth(macs), 15000);
    return () => clearInterval(id);
  }, [commands, refreshWatchHealth]);

  async function sendCommand() {
    const mac = formMac === "__custom__" ? formCustomMac.trim() : formMac.trim();
    if (!mac) { alert("MAC adresi zorunludur."); return; }
    if (!formModule.trim()) { alert("Modül adı zorunludur."); return; }
    setSending(true);
    try {
      const r = await fetch("/api/commands/", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ mac, moduleName: formModule.trim(), param: formParam || null }),
      });
      const j = await r.json();
      if (j.success) {
        setFormModule("");
        setFormParam("");
        refresh();
      } else {
        alert(j.error ?? "Hata oluştu.");
      }
    } finally {
      setSending(false);
    }
  }

  async function deleteCmd(id: number, e: React.MouseEvent) {
    e.stopPropagation();
    e.preventDefault();
    if (!confirm(`Komut #${id} silinsin mi?`)) return;
    try {
      const r = await fetch(`/api/commands/?id=${id}`, { method: "DELETE" });
      const j = await r.json();
      if (j.success) {
        setCommands((prev) => prev.filter((c) => c.id !== id));
      } else {
        alert(j.message ?? j.error ?? "Silinemedi");
      }
    } catch {
      alert("Silme isteği başarısız");
    }
  }

  async function clearStuckCommands() {
    const stuck = commands.filter((c) => c.status === "error" || c.status === "running");
    if (stuck.length === 0) return;
    if (!confirm(`${stuck.length} takılı/hatalı komut silinsin mi?`)) return;
    for (const c of stuck) {
      await fetch(`/api/commands/?id=${c.id}`, { method: "DELETE" });
    }
    refresh();
  }

  function toggleExpand(id: number) {
    setExpanded((prev) => {
      const next = new Set(prev);
      next.has(id) ? next.delete(id) : next.add(id);
      return next;
    });
  }

  const filtered = commands.filter((c) => {
    if (filterStatus !== "tümü" && c.status !== filterStatus) return false;
    if (filterMac && !c.mac.toLowerCase().includes(filterMac.toLowerCase())) return false;
    return true;
  });

  const counts = { pending: 0, running: 0, done: 0, error: 0 };
  commands.forEach((c) => { if (c.status in counts) counts[c.status as keyof typeof counts]++; });

  return (
    <div className="page-wrap">
      <div className="page-header">
        <div>
          <div className="page-title">Uzaktan Komut Kuyruğu</div>
          <div className="page-sub">
            Dashboard'dan VBA istemcilerine modül çalıştırma komutu gönder
          </div>
        </div>
        <button className="btn btn-ghost" onClick={refresh}>↻ Yenile</button>
        {counts.error + counts.running > 0 && (
          <button className="btn btn-danger" style={{ marginLeft: 8 }} onClick={clearStuckCommands}>
            Takılı/hatalı komutları sil
          </button>
        )}
      </div>

      {/* Stats */}
      <div style={{ display: "flex", gap: 10, marginBottom: 20, flexWrap: "wrap" }}>
        {Object.entries(counts).map(([st, cnt]) => {
          const c = STATUS_CFG[st];
          return (
            <div key={st} style={{
              padding: "10px 18px", borderRadius: "var(--radius-sm)",
              background: c.bg, border: `1px solid ${c.color}40`,
              minWidth: 110, textAlign: "center",
            }}>
              <div style={{ fontSize: 22, fontWeight: 700, color: c.color }}>{cnt}</div>
              <div style={{ fontSize: 11, color: c.color, marginTop: 2 }}>{c.label}</div>
            </div>
          );
        })}
      </div>

      {/* Yeni komut formu */}
      <div className="card" style={{ marginBottom: 20 }}>
        <div className="card-header">
          <span className="card-title">🚀 Yeni Komut Gönder</span>
        </div>
        <div style={{ padding: "16px 18px", display: "flex", flexDirection: "column", gap: 14 }}>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            <label style={LS}>
              <span>Hedef MAC</span>
              <select
                className="form-input"
                value={formMac}
                onChange={(e) => setFormMac(e.target.value)}
              >
                {allMacs.map((m) => (
                  <option key={m} value={m}>{m}</option>
                ))}
                <option value="__custom__">— Özel MAC —</option>
              </select>
              {formMac === "__custom__" && (
                <input
                  className="form-input mono"
                  placeholder="XX:XX:XX:XX:XX:XX"
                  value={formCustomMac}
                  onChange={(e) => setFormCustomMac(e.target.value)}
                  style={{ marginTop: 6 }}
                />
              )}
            </label>
            <label style={LS}>
              <span>Modül Adı</span>
              <select
                className="form-input"
                value={formModule}
                onChange={(e) => setFormModule(e.target.value)}
              >
                <option value="">— Seçin —</option>
                {allModuleNames.map((m) => (
                  <option key={m} value={m}>{m}</option>
                ))}
              </select>
            </label>
          </div>
          <label style={LS}>
            <span>Parametre (isteğe bağlı)</span>
            <input
              className="form-input"
              placeholder='Örn: C:\Users\demo  veya  {"key":"val"}'
              value={formParam}
              onChange={(e) => setFormParam(e.target.value)}
            />
          </label>
          <div style={{ display: "flex", justifyContent: "flex-end" }}>
            <button className="btn btn-primary" onClick={sendCommand} disabled={sending}>
              {sending ? "Gönderiliyor…" : "📤 Komutu Kuyruğa Ekle"}
            </button>
          </div>
        </div>
      </div>

      {/* Filtreler */}
      <div style={{ display: "flex", gap: 8, marginBottom: 12, flexWrap: "wrap", alignItems: "center" }}>
        {["tümü", "pending", "running", "done", "error"].map((st) => (
          <button
            key={st}
            onClick={() => setFilterStatus(st)}
            style={{
              padding: "4px 12px", borderRadius: 20, fontSize: 12, cursor: "pointer",
              border: "1px solid var(--border)",
              background: filterStatus === st ? "var(--accent)" : "var(--bg-card)",
              color: filterStatus === st ? "#fff" : "var(--text-muted)",
            }}
          >
            {STATUS_CFG[st]?.label ?? "Tümü"} {st !== "tümü" && `(${counts[st as keyof typeof counts] ?? 0})`}
          </button>
        ))}
        <input
          className="form-input"
          placeholder="MAC filtrele…"
          value={filterMac}
          onChange={(e) => setFilterMac(e.target.value)}
          style={{ maxWidth: 220 }}
        />
        <span style={{ fontSize: 12, color: "var(--text-dim)" }}>
          {filtered.length}/{commands.length}
        </span>
      </div>

      {/* Komut listesi */}
      <div className="card" style={{ padding: 0 }}>
        {filtered.length === 0 ? (
          <div className="empty-state">
            <div className="empty-state-icon">📭</div>
            <div>Komut bulunamadı.</div>
          </div>
        ) : (
          filtered.map((cmd) => {
            const isOpen = expanded.has(cmd.id);
            const runSec = cmd.status === "running" ? elapsedSecSince(cmd.executedAt) : null;
            const isWatch = cmd.moduleName === "WatchFolderServer";
            const health = isWatch ? watchHealth[cmd.mac] : undefined;
            return (
              <div key={cmd.id} style={{ borderBottom: "1px solid var(--border)" }}>
                <div style={{
                  padding: "12px 18px", display: "flex",
                  alignItems: "center", gap: 12, flexWrap: "wrap",
                }}>
                  <button
                    onClick={() => toggleExpand(cmd.id)}
                    style={{
                      background: "none", border: "none", cursor: "pointer",
                      color: "var(--text-dim)", fontSize: 11, padding: 4,
                      transform: isOpen ? "rotate(90deg)" : "none",
                      transition: "transform 0.15s",
                    }}
                  >▶</button>

                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
                      <span style={{ fontSize: 11, color: "var(--text-dim)" }}>#{cmd.id}</span>
                      <span className="mono" style={{ fontSize: 13, fontWeight: 600 }}>
                        {cmd.moduleName}
                      </span>
                      <StatusBadge status={cmd.status} />
                      {runSec != null && runSec >= 30 && (
                        <span style={{
                          fontSize: 10, padding: "2px 8px", borderRadius: 8,
                          background: runSec >= 120 ? "#ef444420" : "#f59e0b20",
                          color: runSec >= 120 ? "#ef4444" : "#f59e0b",
                          fontWeight: 600,
                        }}>
                          {formatDurationSec(runSec)} süredir
                        </span>
                      )}
                    </div>
                    <div style={{ fontSize: 11, color: "var(--text-muted)", marginTop: 2 }}>
                      <span className="mono">{cmd.mac}</span>
                      {cmd.param && <span style={{ marginLeft: 8 }}>param: {cmd.param.slice(0, 40)}</span>}
                    </div>
                    {isWatch && cmd.status === "running" && (
                      <div style={{ fontSize: 11, marginTop: 4 }}>
                        {health?.isAlive ? (
                          <span style={{ color: "var(--green)" }}>
                            ✓ Klasör izleme aktif — son sinyal {timeAgo(health.lastPingAt)}
                          </span>
                        ) : health?.lastPingAt ? (
                          <span style={{ color: "#f59e0b" }}>
                            ⚠ Son sinyal {timeAgo(health.lastPingAt)} —{" "}
                            <Link href="/klasor-izleme" style={{ color: "var(--accent)" }}>Klasör İzleme</Link>
                          </span>
                        ) : (
                          <span style={{ color: "var(--text-dim)" }}>
                            Henüz sunucudan sinyal yok — modül hâlâ başlıyor veya takılı olabilir
                          </span>
                        )}
                        {runSec != null && runSec >= 60 && (
                          <span style={{ display: "block", color: "var(--text-dim)", marginTop: 2 }}>
                            Not: WatchFolderServer saniyeler içinde biter; komut satırı takılı kalsa bile izleme arka planda çalışabilir.
                          </span>
                        )}
                      </div>
                    )}
                  </div>

                  <div style={{ fontSize: 11, color: "var(--text-dim)", textAlign: "right" }}>
                    <div title="Oluşturulma">{formatTR(cmd.createdAt)}</div>
                    {cmd.executedAt && (
                      <div style={{ marginTop: 2 }} title="Çalıştırılma">
                        {cmd.status === "running" ? "▶" : "∆"} {formatTR(cmd.executedAt)}
                        {runSec != null && cmd.status === "running" && (
                          <span style={{ marginLeft: 4, color: runSec >= 120 ? "#ef4444" : "inherit" }}>
                            ({formatDurationSec(runSec)})
                          </span>
                        )}
                      </div>
                    )}
                  </div>

                  <button
                    type="button"
                    className="btn btn-danger"
                    style={{ fontSize: 11, padding: "3px 9px" }}
                    onClick={(e) => deleteCmd(cmd.id, e)}
                  >Sil</button>
                </div>

                {isOpen && (
                  <div style={{
                    background: "var(--bg)", borderTop: "1px solid var(--border)",
                    padding: "14px 18px",
                  }}>
                    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
                      <div>
                        <div style={{ fontSize: 11, fontWeight: 600, color: "var(--text-muted)", marginBottom: 4 }}>
                          KOMUT DETAYI
                        </div>
                        <div style={{ fontSize: 12 }}>
                          <div>ID: {cmd.id}</div>
                          <div>MAC: <span className="mono">{cmd.mac}</span></div>
                          <div>Modül: <span className="mono">{cmd.moduleName}</span></div>
                          {cmd.param && <div>Param: <span className="mono">{cmd.param}</span></div>}
                          <div>Oluşturan: {cmd.createdBy}</div>
                        </div>
                      </div>
                      {(cmd.result || cmd.errorMsg) && (
                        <div>
                          <div style={{
                            fontSize: 11, fontWeight: 600,
                            color: cmd.status === "error" ? "#ef4444" : "var(--text-muted)",
                            marginBottom: 4
                          }}>
                            {cmd.status === "error" ? "HATA" : "SONUÇ"}
                          </div>
                          <pre style={{
                            fontSize: 11, fontFamily: "var(--font-geist-mono, monospace)",
                            color: cmd.status === "error" ? "#ef4444" : "var(--text)",
                            background: "var(--code-bg)", padding: "8px 10px",
                            borderRadius: "var(--radius-sm)", overflowX: "auto",
                            maxHeight: 150, whiteSpace: "pre-wrap", wordBreak: "break-all",
                          }}>
                            {cmd.result ?? cmd.errorMsg}
                          </pre>
                        </div>
                      )}
                      {cmd.status === "done" && (
                        <div>
                          <div style={{ fontSize: 11, fontWeight: 600, color: "var(--text-muted)", marginBottom: 4 }}>
                            MODÜL ÇIKTISI
                          </div>
                          <div style={{ fontSize: 12 }}>
                            Veriler Excel sayfasına yazılır ve{" "}
                            <a href="/modul-ciktilari/" style={{ color: "var(--accent)" }}>
                              Modül Çıktıları
                            </a>{" "}
                            sayfasında görünür.
                          </div>
                        </div>
                      )}
                    </div>
                  </div>
                )}
              </div>
            );
          })
        )}
      </div>

      <div style={{ marginTop: 16, padding: "12px 16px", background: "var(--bg-card)", borderRadius: "var(--radius-sm)", border: "1px solid var(--border)", fontSize: 12, color: "var(--text-muted)" }}>
        <strong>Komut kuyruğu:</strong> Excel açıkken{" "}
        <code style={{ fontFamily: "var(--font-geist-mono, monospace)", fontSize: 11 }}>
          InstallCommandQueue
        </code>{" "}
        modülü gizli <code>TeklifPollHost</code> workbook ile ~60 sn&apos;de bir{" "}
        <code style={{ fontFamily: "var(--font-geist-mono, monospace)", fontSize: 11 }}>
          GET /api/commands/pending/&#123;mac&#125;
        </code>{" "}
        sorgular. TeklifAgent.exe yalnızca heartbeat gönderir.         Modül sonuçları{" "}
        <a href="/modul-ciktilari/" style={{ color: "var(--accent)" }}>Modül Çıktıları</a>{" "}
        sayfasında listelenir (komut satırında değil).{" "}
        <strong>WatchFolderServer</strong> gibi arka plan modülleri saniyeler içinde biter; komut satırı
        &quot;Çalışıyor&quot; takılı kalırsa{" "}
        <Link href="/klasor-izleme" style={{ color: "var(--accent)" }}>Klasör İzleme</Link>
        {" "}sayfasında son sinyal zamanına bakın (90 sn içindeyse aktif). 5 dk sonra otomatik hata olur.
      </div>
    </div>
  );
}

const LS: React.CSSProperties = {
  display: "flex", flexDirection: "column", gap: 6,
  fontSize: 13, fontWeight: 500, color: "var(--text-muted)",
};
