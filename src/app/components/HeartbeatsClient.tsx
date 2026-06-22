"use client";

import { useEffect, useState, useRef } from "react";
import type { HeartbeatRow } from "@/lib/db";

/* ── Durum hesabı ──────────────────────────────────── */
type Status = "online" | "idle" | "offline";

function getStatus(lastSeen: string): Status {
  const diffMs = Date.now() - new Date(lastSeen).getTime();
  const diffMin = diffMs / 60_000;
  if (diffMin < 2)  return "online";
  if (diffMin < 10) return "idle";
  return "offline";
}

const STATUS_STYLE: Record<Status, { dot: string; label: string; color: string }> = {
  online:  { dot: "#22c55e", label: "Çevrimiçi", color: "#16a34a" },
  idle:    { dot: "#f59e0b", label: "Boşta",     color: "#b45309" },
  offline: { dot: "#ef4444", label: "Çevrimdışı",color: "#b91c1c" },
};

/* ── Zaman göstergesi ──────────────────────────────── */
function TimeAgo({ iso }: { iso: string }) {
  const [text, setText] = useState("");

  useEffect(() => {
    function calc() {
      const diffMs  = Date.now() - new Date(iso).getTime();
      const diffSec = Math.floor(diffMs / 1000);
      if (diffSec < 60)  return `${diffSec} sn önce`;
      const diffMin = Math.floor(diffSec / 60);
      if (diffMin < 60)  return `${diffMin} dk önce`;
      const diffH   = Math.floor(diffMin / 60);
      if (diffH   < 24)  return `${diffH} sa önce`;
      return `${Math.floor(diffH / 24)} gün önce`;
    }
    setText(calc());
    const t = setInterval(() => setText(calc()), 10_000);
    return () => clearInterval(t);
  }, [iso]);

  const local = new Date(iso).toLocaleString("tr-TR", {
    timeZone: "Europe/Istanbul",
    day:    "2-digit", month: "2-digit", year: "numeric",
    hour:   "2-digit", minute: "2-digit", second: "2-digit",
  });

  return (
    <span title={local} style={{ cursor: "help" }}>{text}</span>
  );
}

/* ── Ana bileşen ───────────────────────────────────── */
export default function HeartbeatsClient({ initial }: { initial: HeartbeatRow[] }) {
  const [rows, setRows]             = useState<HeartbeatRow[]>(initial);
  const [lastRefresh, setLastRefresh] = useState<Date>(new Date());
  const [countdown, setCountdown]   = useState(30);
  const timerRef = useRef<ReturnType<typeof setInterval> | null>(null);

  function refresh() {
    fetch("/api/heartbeat")
      .then((r) => r.json())
      .then((j) => {
        if (j.success) { setRows(j.data); setLastRefresh(new Date()); }
      });
  }

  /* 30 sn otomatik yenileme */
  useEffect(() => {
    setCountdown(30);
    timerRef.current = setInterval(() => {
      setCountdown((c) => {
        if (c <= 1) { refresh(); return 30; }
        return c - 1;
      });
    }, 1000);
    return () => { if (timerRef.current) clearInterval(timerRef.current); };
  }, []);

  const online  = rows.filter((r) => getStatus(r.last_seen) === "online").length;
  const idle    = rows.filter((r) => getStatus(r.last_seen) === "idle").length;
  const offline = rows.filter((r) => getStatus(r.last_seen) === "offline").length;

  return (
    <div className="page-wrap">
      {/* Başlık */}
      <div className="page-header">
        <div>
          <div className="page-title">Cihaz Nabız İzleme</div>
          <div className="page-sub">
            Excel istemcilerinin son aktiflik bilgileri — her {" "}
            <span className="mono">HeartbeatPing</span> çalışmasında güncellenir
          </div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          <span style={{ fontSize: 12, color: "var(--text-muted)" }}>
            {countdown}s sonra yenilenir
          </span>
          <button className="btn btn-ghost" onClick={() => { refresh(); setCountdown(30); }}>
            ↻ Yenile
          </button>
        </div>
      </div>

      {/* Özet kartlar */}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 12, marginBottom: 20 }}>
        {([ ["online", online], ["idle", idle], ["offline", offline] ] as [Status, number][]).map(([st, n]) => (
          <div key={st} className="card" style={{ padding: "14px 18px" }}>
            <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
              <span style={{
                width: 10, height: 10, borderRadius: "50%",
                background: STATUS_STYLE[st].dot, flexShrink: 0,
              }} />
              <span style={{ fontSize: 13, color: "var(--text-muted)", flex: 1 }}>
                {STATUS_STYLE[st].label}
              </span>
              <span style={{ fontSize: 22, fontWeight: 700, color: STATUS_STYLE[st].color }}>
                {n}
              </span>
            </div>
          </div>
        ))}
      </div>

      {/* Tablo */}
      {rows.length === 0 ? (
        <div className="card">
          <div className="empty-state">
            <div className="empty-state-icon">📡</div>
            <div>Henüz hiç heartbeat gelmedi.</div>
            <div style={{ fontSize: 12, color: "var(--text-dim)", marginTop: 6 }}>
              Excel&apos;de <span className="mono">HeartbeatPing</span> modülünü çalıştırın.
            </div>
          </div>
        </div>
      ) : (
        <div className="card" style={{ overflow: "hidden" }}>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead>
              <tr style={{ borderBottom: "1px solid var(--border)" }}>
                {["Durum", "MAC Adresi", "Bilgisayar", "Kullanıcı", "Excel", "IP", "Son Görülme"].map((h) => (
                  <th key={h} style={TH}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {rows.map((row) => {
                const st = getStatus(row.last_seen);
                const s  = STATUS_STYLE[st];
                return (
                  <tr key={row.mac} style={{ borderBottom: "1px solid var(--border)" }}
                    onMouseEnter={(e) => (e.currentTarget.style.background = "var(--bg)")}
                    onMouseLeave={(e) => (e.currentTarget.style.background = "")}
                  >
                    {/* Durum */}
                    <td style={TD}>
                      <span style={{ display: "flex", alignItems: "center", gap: 6 }}>
                        <span style={{
                          width: 8, height: 8, borderRadius: "50%",
                          background: s.dot, flexShrink: 0,
                          boxShadow: st === "online" ? `0 0 6px ${s.dot}` : "none",
                        }} />
                        <span style={{ fontSize: 12, color: s.color, fontWeight: 600 }}>
                          {s.label}
                        </span>
                      </span>
                    </td>

                    {/* MAC */}
                    <td style={TD}>
                      <span className="mono" style={{ fontSize: 12 }}>{row.mac}</span>
                    </td>

                    {/* Hostname */}
                    <td style={TD}>
                      <span style={{ fontSize: 13 }}>{row.hostname ?? <Dim />}</span>
                    </td>

                    {/* Kullanıcı */}
                    <td style={TD}>
                      <span style={{ fontSize: 13 }}>{row.user_name ?? <Dim />}</span>
                    </td>

                    {/* Excel versiyon */}
                    <td style={TD}>
                      <span className="mono" style={{ fontSize: 12 }}>
                        {row.excel_version ? `v${row.excel_version}` : <Dim />}
                      </span>
                    </td>

                    {/* IP */}
                    <td style={TD}>
                      <span className="mono" style={{ fontSize: 12 }}>{row.ip_address ?? <Dim />}</span>
                    </td>

                    {/* Son görülme */}
                    <td style={{ ...TD, color: "var(--text-muted)", fontSize: 12 }}>
                      <TimeAgo iso={row.last_seen} />
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}

      <div style={{ marginTop: 12, fontSize: 11, color: "var(--text-dim)", textAlign: "right" }}>
        Son güncelleme: {lastRefresh.toLocaleTimeString("tr-TR")}
        &nbsp;·&nbsp;
        {rows.length} cihaz kayıtlı
      </div>
    </div>
  );
}

/* ── Style sabitleri ───────────────────────────────── */
const TH: React.CSSProperties = {
  padding: "10px 14px",
  textAlign: "left",
  fontSize: 11,
  fontWeight: 600,
  textTransform: "uppercase",
  letterSpacing: "0.05em",
  color: "var(--text-muted)",
};

const TD: React.CSSProperties = {
  padding: "11px 14px",
  verticalAlign: "middle",
};

function Dim() {
  return <span style={{ color: "var(--text-dim)" }}>—</span>;
}
