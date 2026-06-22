export const dynamic = "force-dynamic";

import Link from "next/link";
import {
  listLicenses,
  listLogs,
  listHeartbeats,
  listDeviceSnapshots,
  listClientCommands,
  listModuleOutputs,
} from "@/lib/db";
import { formatTR, getHeartbeatStatus, timeAgo } from "@/lib/date-utils";
import { listFirmAutoModules } from "@/lib/firm-auto-modules";
import { listModules } from "@/lib/modules";
import { loadProposalsSummary } from "@/lib/proposals";
import { getApiCatalogStats } from "@/lib/api-catalog";
import Refresher from "./components/Refresher";

const QUICK_LINKS = [
  { href: "/komutlar",        icon: "🎮", label: "Uzak Komutlar"    },
  { href: "/heartbeats",      icon: "📡", label: "Nabız İzleme"     },
  { href: "/cihazlar",        icon: "🖥", label: "Cihaz Bilgileri"  },
  { href: "/modul-ciktilari", icon: "📤", label: "Modül Çıktıları"  },
  { href: "/moduller",        icon: "📦", label: "Uzak Modüller"    },
  { href: "/firma-modulleri", icon: "⚡", label: "Oto. Modüller"    },
  { href: "/oneriler",        icon: "💡", label: "Modül Önerileri"  },
  { href: "/api-referans",    icon: "🔌", label: "API Referans"     },
];

function LicenseBadge({ value }: { value: string }) {
  const active = ["true", "1", "active", "evet"].includes(value.toLowerCase());
  return (
    <span className={`badge ${active ? "badge-green" : "badge-red"}`}>
      <span className="badge-dot" />{active ? "Aktif" : "Pasif"}
    </span>
  );
}

function CmdBadge({ status }: { status: string }) {
  const map: Record<string, string> = {
    pending: "badge-yellow",
    running: "badge-blue",
    done: "badge-green",
    error: "badge-red",
  };
  const labels: Record<string, string> = {
    pending: "Bekliyor",
    running: "Çalışıyor",
    done: "Tamam",
    error: "Hata",
  };
  return <span className={`badge ${map[status] ?? "badge-blue"}`}>{labels[status] ?? status}</span>;
}

function HbDot({ lastSeen }: { lastSeen: string }) {
  const st = getHeartbeatStatus(lastSeen);
  const color = st === "online" ? "var(--green)" : st === "idle" ? "var(--yellow)" : "var(--red)";
  return (
    <span style={{ width: 8, height: 8, borderRadius: "50%", background: color, display: "inline-block" }} />
  );
}

export default async function Dashboard() {
  const [
    licenses,
    modules,
    firmAutoModules,
    logs,
    heartbeats,
    snapshots,
    commands,
    outputs,
    proposals,
  ] = await Promise.all([
    listLicenses(),
    listModules(),
    listFirmAutoModules(),
    listLogs(8),
    listHeartbeats(),
    listDeviceSnapshots(),
    listClientCommands({ limit: 200 }),
    listModuleOutputs({ limit: 1 }),
    Promise.resolve(loadProposalsSummary()),
  ]);

  const apiStats = getApiCatalogStats();

  const activeLicenses = licenses.filter((l) =>
    ["true", "1", "active", "evet"].includes(l.license.toLowerCase()),
  );

  const onlineCount = heartbeats.filter((h) => getHeartbeatStatus(h.last_seen) === "online").length;
  const idleCount = heartbeats.filter((h) => getHeartbeatStatus(h.last_seen) === "idle").length;
  const offlineCount = heartbeats.length - onlineCount - idleCount;

  const cmdPending = commands.filter((c) => c.status === "pending" || c.status === "running").length;
  const cmdError = commands.filter((c) => c.status === "error").length;

  const categoryCounts: Record<string, number> = {};
  for (const m of modules) {
    const cat = (m as { category?: string }).category ?? "genel";
    categoryCounts[cat] = (categoryCounts[cat] ?? 0) + 1;
  }
  const topCategories = Object.entries(categoryCounts)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 6);

  return (
    <div className="page-wrap">
      <div className="page-header">
        <div>
          <div className="page-title">Dashboard</div>
          <div className="page-sub">
            Teklif Sunucu — {apiStats.total} API endpoint · {proposals.stats.inDb} uzak modül
          </div>
        </div>
        <Refresher />
      </div>

      {/* Ana metrikler */}
      <div className="stats-grid" style={{ marginBottom: 16 }}>
        <div className="stat-card">
          <div className="stat-label">Lisanslar</div>
          <div className="stat-value">{licenses.length}</div>
          <div className="stat-hint">{activeLicenses.length} aktif</div>
        </div>
        <div className="stat-card">
          <div className="stat-label">Çevrimiçi</div>
          <div className="stat-value green">{onlineCount}</div>
          <div className="stat-hint">{idleCount} boşta · {offlineCount} kapalı</div>
        </div>
        <div className="stat-card">
          <div className="stat-label">Bekleyen Komut</div>
          <div className="stat-value" style={{ color: cmdPending > 0 ? "var(--yellow)" : undefined }}>
            {cmdPending}
          </div>
          <div className="stat-hint">{cmdError > 0 ? `${cmdError} hatalı` : "Kuyruk temiz"}</div>
        </div>
        <div className="stat-card">
          <div className="stat-label">Modül Önerileri</div>
          <div className="stat-value accent">{proposals.stats.done}</div>
          <div className="stat-hint">
            {proposals.stats.missingFromDb > 0
              ? `${proposals.stats.missingFromDb} DB eksik`
              : `${proposals.stats.planned} planlı`}
          </div>
        </div>
      </div>

      {/* Hızlı erişim */}
      <div style={{
        display: "grid",
        gridTemplateColumns: "repeat(auto-fill, minmax(140px, 1fr))",
        gap: 10,
        marginBottom: 20,
      }}>
        {QUICK_LINKS.map((l) => (
          <Link
            key={l.href}
            href={l.href}
            className="card"
            style={{
              padding: "12px 14px",
              display: "flex",
              alignItems: "center",
              gap: 10,
              transition: "border-color 0.15s",
            }}
          >
            <span style={{ fontSize: 18 }}>{l.icon}</span>
            <span style={{ fontSize: 12, fontWeight: 600 }}>{l.label}</span>
          </Link>
        ))}
      </div>

      <div className="grid-2" style={{ marginBottom: 16 }}>
        {/* Nabız */}
        <div className="card">
          <div className="card-header">
            <span className="card-title">📡 Son Nabız</span>
            <Link href="/heartbeats" style={{ fontSize: 12, color: "var(--accent)" }}>Tümü →</Link>
          </div>
          {heartbeats.length === 0 ? (
            <div className="empty-state" style={{ padding: 24 }}><div>Henüz heartbeat yok</div></div>
          ) : (
            <div className="table-wrap">
              <table className="data-table">
                <thead><tr><th></th><th>MAC</th><th>PC</th><th>Son</th></tr></thead>
                <tbody>
                  {heartbeats.slice(0, 5).map((h) => (
                    <tr key={h.mac}>
                      <td><HbDot lastSeen={h.last_seen} /></td>
                      <td><span className="mono" style={{ fontSize: 11 }}>{h.mac}</span></td>
                      <td style={{ fontSize: 13 }}>{h.hostname ?? "—"}</td>
                      <td style={{ fontSize: 12, color: "var(--text-muted)" }}>{timeAgo(h.last_seen)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>

        {/* Komutlar */}
        <div className="card">
          <div className="card-header">
            <span className="card-title">🎮 Son Komutlar</span>
            <Link href="/komutlar" style={{ fontSize: 12, color: "var(--accent)" }}>Tümü →</Link>
          </div>
          {commands.length === 0 ? (
            <div className="empty-state" style={{ padding: 24 }}><div>Komut yok</div></div>
          ) : (
            <div className="table-wrap">
              <table className="data-table">
                <thead><tr><th>Modül</th><th>MAC</th><th>Durum</th></tr></thead>
                <tbody>
                  {commands.slice(0, 5).map((c) => (
                    <tr key={c.id}>
                      <td><span className="mono" style={{ fontSize: 12 }}>{c.moduleName}</span></td>
                      <td><span className="mono" style={{ fontSize: 11 }}>{c.mac}</span></td>
                      <td><CmdBadge status={c.status} /></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>
      </div>

      <div className="grid-2" style={{ marginBottom: 16 }}>
        {/* Lisanslar */}
        <div className="card">
          <div className="card-header">
            <span className="card-title">🔑 Son Lisanslar</span>
            <Link href="/lisanslar" style={{ fontSize: 12, color: "var(--accent)" }}>Tümü →</Link>
          </div>
          <div className="table-wrap">
            <table className="data-table">
              <thead><tr><th>MAC</th><th>Firma</th><th>Lisans</th></tr></thead>
              <tbody>
                {licenses.slice(0, 5).map((item) => (
                  <tr key={item.macAdresi}>
                    <td><span className="mono" style={{ fontSize: 11 }}>{item.macAdresi}</span></td>
                    <td style={{ fontSize: 13 }}>{item.firmaAdi ?? "—"}</td>
                    <td><LicenseBadge value={item.license} /></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>

        {/* Loglar */}
        <div className="card">
          <div className="card-header">
            <span className="card-title">📋 Son Aktivite</span>
            <Link href="/loglar" style={{ fontSize: 12, color: "var(--accent)" }}>Tümü →</Link>
          </div>
          <div className="table-wrap">
            <table className="data-table">
              <thead><tr><th>Olay</th><th>MAC</th><th>Zaman</th></tr></thead>
              <tbody>
                {logs.map((log) => (
                  <tr key={log.id}>
                    <td style={{ fontSize: 12 }}>{log.eventType}</td>
                    <td><span className="mono" style={{ fontSize: 11 }}>{log.macAdresi ?? "—"}</span></td>
                    <td style={{ fontSize: 11, color: "var(--text-muted)" }}>{formatTR(log.createdAt)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      </div>

      <div className="grid-2">
        {/* Modül kategorileri */}
        <div className="card">
          <div className="card-header">
            <span className="card-title">📦 Modül Kategorileri</span>
            <Link href="/moduller" style={{ fontSize: 12, color: "var(--accent)" }}>{modules.length} modül →</Link>
          </div>
          <div style={{ padding: "14px 18px", display: "flex", flexDirection: "column", gap: 8 }}>
            {topCategories.map(([cat, n]) => (
              <div key={cat} style={{ display: "flex", alignItems: "center", gap: 10 }}>
                <span style={{ width: 72, fontSize: 11, color: "var(--text-muted)", textTransform: "capitalize" }}>{cat}</span>
                <div style={{ flex: 1, height: 6, background: "var(--border)", borderRadius: 3, overflow: "hidden" }}>
                  <div style={{
                    width: `${Math.round((n / modules.length) * 100)}%`,
                    height: "100%",
                    background: "var(--accent)",
                    borderRadius: 3,
                  }} />
                </div>
                <span style={{ fontSize: 12, fontWeight: 600, minWidth: 24, textAlign: "right" }}>{n}</span>
              </div>
            ))}
          </div>
        </div>

        {/* Sistem özeti */}
        <div className="card">
          <div className="card-header">
            <span className="card-title">⚙ Sistem Özeti</span>
            <Link href="/oneriler" style={{ fontSize: 12, color: "var(--accent)" }}>Öneriler →</Link>
          </div>
          <div style={{ padding: "14px 18px", fontSize: 13, color: "var(--text-muted)", display: "flex", flexDirection: "column", gap: 10 }}>
            <div style={{ display: "flex", justifyContent: "space-between" }}>
              <span>Cihaz snapshot</span>
              <strong style={{ color: "var(--text)" }}>{snapshots.length}</strong>
            </div>
            <div style={{ display: "flex", justifyContent: "space-between" }}>
              <span>Modül çıktısı (son kayıtlar)</span>
              <strong style={{ color: "var(--text)" }}>{outputs.length > 0 ? "✓ aktif" : "—"}</strong>
            </div>
            <div style={{ display: "flex", justifyContent: "space-between" }}>
              <span>Firma oto-modül tanımı</span>
              <strong style={{ color: "var(--text)" }}>{firmAutoModules.length}</strong>
            </div>
            <div style={{ display: "flex", justifyContent: "space-between" }}>
              <span>API endpoint</span>
              <strong style={{ color: "var(--text)" }}>{apiStats.total}</strong>
            </div>
            <div style={{ display: "flex", justifyContent: "space-between" }}>
              <span>Öneri → DB uyumu</span>
              <strong style={{ color: proposals.stats.missingFromDb > 0 ? "var(--yellow)" : "var(--green)" }}>
                {proposals.stats.missingFromDb === 0 ? "Tam" : `${proposals.stats.missingFromDb} eksik`}
              </strong>
            </div>
            {firmAutoModules.length > 0 && (
              <div style={{ marginTop: 6, paddingTop: 10, borderTop: "1px solid var(--border)", fontSize: 12 }}>
                <div style={{ fontWeight: 600, marginBottom: 6, color: "var(--text)" }}>Global açılış zinciri</div>
                {(firmAutoModules.find((f) => f.firmaAdi === "*")?.onExcelOpen.modules ?? [])
                  .sort((a, b) => a.order - b.order)
                  .map((m) => m.methodName + (m.runOnce ? " (1×)" : ""))
                  .join(" → ") || "—"}
              </div>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}
