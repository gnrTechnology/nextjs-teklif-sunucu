"use client";

import { useCallback, useEffect, useState } from "react";
import type { FolderWatchEvent, FolderWatchHealth } from "@/lib/types";
import { formatTR, timeAgo, formatDurationSec, elapsedSecSince } from "@/lib/date-utils";
import Link from "next/link";

const EVENT_STYLE: Record<string, { label: string; color: string }> = {
  created: { label: "Oluşturuldu", color: "var(--green)" },
  deleted: { label: "Silindi", color: "var(--red)" },
  modified: { label: "Değişti", color: "var(--yellow)" },
  started: { label: "Başlatıldı", color: "var(--accent)" },
  scan: { label: "Canlı ping", color: "var(--text-muted)" },
};

export default function KlasorIzlemeClient({
  initial,
  heartbeats,
}: {
  initial: FolderWatchEvent[];
  heartbeats: { mac: string; hostname: string | null }[];
}) {
  const [events, setEvents] = useState(initial);
  const [mac, setMac] = useState(heartbeats[0]?.mac ?? "");
  const [health, setHealth] = useState<FolderWatchHealth | null>(null);
  const [msg, setMsg] = useState("");
  const [busy, setBusy] = useState(false);
  const [folderPath, setFolderPath] = useState("C:\\Users\\onurm\\Desktop\\");
  const [hideScan, setHideScan] = useState(true);

  const refresh = useCallback(async () => {
    const q = mac ? `?mac=${encodeURIComponent(mac)}&limit=200` : "?limit=200";
    const r = await fetch(`/api/folder-watch${q}`);
    const j = await r.json();
    if (j.success) setEvents(j.data);
    if (mac) {
      const hr = await fetch(`/api/folder-watch/?health=1&mac=${encodeURIComponent(mac)}`);
      const hj = await hr.json();
      if (hj.success) setHealth(hj.data?.[0] ?? null);
    } else {
      setHealth(null);
    }
  }, [mac]);

  useEffect(() => {
    refresh();
    const t = setInterval(refresh, 30000);
    return () => clearInterval(t);
  }, [refresh]);

  async function startWatchOnClient() {
    if (!mac) {
      setMsg("MAC seçin.");
      return;
    }
    setBusy(true);
    setMsg("");
    try {
      let path = folderPath.trim() || "C:\\";
      if (!path.endsWith("\\")) path += "\\";
      const param = JSON.stringify({ folderPath: path, intervalSec: 30 });
      const r = await fetch("/api/commands/", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          mac,
          moduleName: "WatchFolderServer",
          param,
          createdBy: "dashboard/klasor-izleme",
        }),
      });
      const j = await r.json();
      setMsg(j.success ? "✓ WatchFolderServer komutu kuyruğa eklendi." : `✗ ${j.error ?? "Hata"}`);
    } catch (e) {
      setMsg(`✗ ${String(e)}`);
    }
    setBusy(false);
  }

  const fileEvents = events.filter((e) => e.eventType !== "scan" && e.eventType !== "started");
  const visibleEvents = hideScan ? events.filter((e) => e.eventType !== "scan") : events;

  const byType = fileEvents.reduce(
    (acc, e) => {
      acc[e.eventType] = (acc[e.eventType] ?? 0) + 1;
      return acc;
    },
    {} as Record<string, number>,
  );

  return (
    <div className="page-wrap">
      <div className="page-header">
        <div>
          <div className="page-title">Klasör İzleme</div>
          <div className="page-sub">
            Herhangi bir klasör (üst seviye) — değişiklikler sunucuya akar
          </div>
        </div>
        <Link href="/loglar" className="btn btn-ghost">Tüm loglar →</Link>
      </div>

      {mac && (
        <div className="card" style={{
          marginBottom: 16, padding: "14px 18px",
          borderColor: health?.isAlive ? "var(--green)" : health?.lastPingAt ? "#f59e0b" : "var(--border)",
        }}>
          <div style={{ fontWeight: 600, marginBottom: 6 }}>İzleme Durumu</div>
          {!health?.lastPingAt ? (
            <div style={{ fontSize: 13, color: "var(--text-muted)" }}>
              Bu MAC için henüz klasör sinyali yok. WatchFolderServer komutu gönderin ve Excel açık olsun.
            </div>
          ) : (
            <div style={{ fontSize: 13 }}>
              <span style={{
                color: health.isAlive ? "var(--green)" : "#f59e0b",
                fontWeight: 600,
              }}>
                {health.isAlive ? "● Aktif" : "● Sessiz"}
              </span>
              <span style={{ color: "var(--text-muted)", marginLeft: 10 }}>
                Son sinyal: {timeAgo(health.lastPingAt)} ({formatTR(health.lastPingAt)})
              </span>
              {health.lastPingAt && elapsedSecSince(health.lastPingAt) != null && (
                <span style={{ color: "var(--text-dim)", marginLeft: 8 }}>
                  — {formatDurationSec(elapsedSecSince(health.lastPingAt)!)} önce
                </span>
              )}
              {health.folderPath && (
                <div style={{ fontSize: 12, color: "var(--text-dim)", marginTop: 4 }}>
                  Klasör: <span className="mono">{health.folderPath}</span>
                  {health.lastEventType && ` · son olay: ${health.lastEventType}`}
                </div>
              )}
              {!health.isAlive && health.lastPingAt && (
                <div style={{ fontSize: 12, color: "#f59e0b", marginTop: 6 }}>
                  90 sn içinde yeni sinyal gelmediyse izleme durmuş olabilir (Excel kapalı veya TeklifPollHost tick çalışmıyor).
                </div>
              )}
            </div>
          )}
        </div>
      )}

      <div className="stats-grid" style={{ marginBottom: 16 }}>
        <div className="stat-card">
          <div className="stat-label">Dosya Olayı</div>
          <div className="stat-value">{fileEvents.length}</div>
        </div>
        <div className="stat-card">
          <div className="stat-label">Yeni Dosya</div>
          <div className="stat-value green">{byType.created ?? 0}</div>
        </div>
        <div className="stat-card">
          <div className="stat-label">Değişen</div>
          <div className="stat-value" style={{ color: "var(--yellow)" }}>{byType.modified ?? 0}</div>
        </div>
        <div className="stat-card">
          <div className="stat-label">Silinen</div>
          <div className="stat-value" style={{ color: "var(--red)" }}>{byType.deleted ?? 0}</div>
        </div>
      </div>

      <div className="card" style={{ marginBottom: 16, padding: "16px 18px" }}>
        <div style={{ fontWeight: 600, marginBottom: 12 }}>İzlemeyi Başlat (istemci)</div>
        <div style={{ display: "flex", flexWrap: "wrap", gap: 10, alignItems: "center" }}>
          <select className="form-input" value={mac} onChange={(e) => setMac(e.target.value)} style={{ minWidth: 220 }}>
            <option value="">MAC seçin…</option>
            {heartbeats.map((h) => (
              <option key={h.mac} value={h.mac}>
                {h.mac} {h.hostname ? `(${h.hostname})` : ""}
              </option>
            ))}
          </select>
          <input
            className="form-input mono"
            value={folderPath}
            onChange={(e) => setFolderPath(e.target.value)}
            placeholder="C:\Users\...\Desktop\"
            style={{ minWidth: 280, flex: 1 }}
            title="İzlenecek klasör yolu"
          />
          <button className="btn btn-primary" disabled={busy || !mac} onClick={startWatchOnClient}>
            WatchFolderServer gönder
          </button>
          <button className="btn btn-ghost" onClick={refresh}>↻ Yenile</button>
        </div>
        {msg && <div style={{ marginTop: 10, fontSize: 12, color: "var(--text-muted)" }}>{msg}</div>}
        <div style={{ marginTop: 12, fontSize: 12, color: "var(--text-dim)", lineHeight: 1.55 }}>
          Excel açık + komut kuyruğu aktif olmalı. İzleme <strong>yalnızca seçilen klasörün içindeki</strong> dosya ve alt klasör adlarına bakar;
          alt klasörlerin içi taranmaz (ör. <code>C:\</code> seçiliyken Desktop’taki dosya değişiklikleri görünmez).
          <br />
          Tarama her ~30 sn; anlık görüntü artık dosyada saklanır (registry 255 karakter sınırı kaldırıldı).
          <code>C:\</code> kökü için yalnızca doğrudan alt öğeler izlenir; binlerce sistem dosyası ve alt ağaç kapsam dışıdır.
          Güvenilir test için <code>Desktop</code> gibi küçük bir klasör yolunu yazıp komutu yeniden gönderin.
        </div>
      </div>

      <div className="card">
        <div className="card-header" style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <span className="card-title">Klasör Olayları</span>
          <label style={{ fontSize: 12, color: "var(--text-muted)", display: "flex", alignItems: "center", gap: 6 }}>
            <input type="checkbox" checked={hideScan} onChange={(e) => setHideScan(e.target.checked)} />
            Canlı ping gizle
          </label>
        </div>
        {visibleEvents.length === 0 ? (
          <div className="empty-state">
            <div className="empty-state-icon">📁</div>
            <div>Henüz olay yok. WatchFolderServer komutunu gönderin.</div>
          </div>
        ) : (
          <div className="table-wrap">
            <table className="data-table">
              <thead>
                <tr>
                  <th>Zaman</th>
                  <th>Olay</th>
                  <th>Dosya</th>
                  <th>Klasör</th>
                  <th>MAC / PC</th>
                  <th>Detay</th>
                </tr>
              </thead>
              <tbody>
                {visibleEvents.map((ev) => {
                  const isDir = ev.detail?.includes("klasor") ?? false;
                  const st = EVENT_STYLE[ev.eventType] ?? { label: ev.eventType, color: "var(--text-muted)" };
                  const label = ev.eventType === "created" && isDir ? "Yeni Klasör" : st.label;
                  return (
                    <tr key={ev.id}>
                      <td style={{ fontSize: 12, color: "var(--text-muted)", whiteSpace: "nowrap" }}>
                        <span title={formatTR(ev.createdAt)}>{timeAgo(ev.createdAt)}</span>
                      </td>
                      <td>
                        <span style={{ color: st.color, fontWeight: 600, fontSize: 12 }}>{label}</span>
                      </td>
                      <td className="mono" style={{ fontSize: 12 }}>{ev.fileName ?? "—"}</td>
                      <td className="mono" style={{ fontSize: 11 }}>{ev.folderPath}</td>
                      <td style={{ fontSize: 12 }}>
                        <div className="mono">{ev.mac}</div>
                        {ev.hostname && <div style={{ color: "var(--text-dim)", fontSize: 11 }}>{ev.hostname}</div>}
                      </td>
                      <td style={{ fontSize: 12, color: "var(--text-muted)" }}>{ev.detail ?? "—"}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
}
