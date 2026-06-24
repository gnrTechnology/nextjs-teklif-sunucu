"use client";

import Link from "next/link";
import { useState } from "react";
import { Radio, Terminal, KeyRound, Package } from "lucide-react";
import type { ClientCommand, HeartbeatRow } from "@/lib/db";
import type { LicenseRecord, ModuleRecord } from "@/lib/types";
import { formatTR, getHeartbeatStatus, timeAgo } from "@/lib/date-utils";
import PageHeader from "./ui/PageHeader";
import StatCard from "./ui/StatCard";
import AlertBanner, { type AlertItem } from "./ui/AlertBanner";
import DataTable from "./ui/DataTable";
import {
  CommandStatusBadge,
  HeartbeatDot,
  LicenseBadge,
} from "@/lib/status-badges";

type LogRow = {
  id: number;
  eventType: string;
  macAdresi?: string | null;
  createdAt: string;
};

type TopModule = Pick<ModuleRecord, "methodName" | "runCount">;

type GeoRow = { mac: string; hostname?: string | null; ip: string; country: string; city?: string };

export default function DashboardClient({
  apiEndpointCount,
  moduleCount,
  activeModuleCount,
  onlineCount,
  idleCount,
  offlineCount,
  cmdPending,
  cmdError,
  activeLicenseCount,
  licenseCount,
  alerts,
  heartbeats,
  commands,
  licenses,
  logs,
  topModules,
  snapshotCount,
  outputCount,
  firmAutoModuleCount,
  globalChain,
  cmdErrorRecent,
  allMacs,
  allModuleNames,
  geoLocations = [],
}: {
  apiEndpointCount: number;
  moduleCount: number;
  activeModuleCount: number;
  onlineCount: number;
  idleCount: number;
  offlineCount: number;
  cmdPending: number;
  cmdError: number;
  activeLicenseCount: number;
  licenseCount: number;
  alerts: AlertItem[];
  heartbeats: HeartbeatRow[];
  commands: ClientCommand[];
  licenses: LicenseRecord[];
  logs: LogRow[];
  topModules: TopModule[];
  snapshotCount: number;
  outputCount: number;
  firmAutoModuleCount: number;
  globalChain: string;
  cmdErrorRecent: ClientCommand[];
  allMacs: string[];
  allModuleNames: string[];
  geoLocations?: GeoRow[];
}) {
  const [qcMac, setQcMac] = useState(allMacs[0] ?? "");
  const [qcModule, setQcModule] = useState("CaptureScreenshot");
  const [qcSending, setQcSending] = useState(false);

  async function sendQuickCommand() {
    if (!qcMac || !qcModule) return;
    setQcSending(true);
    try {
      const r = await fetch("/api/commands/", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ mac: qcMac, moduleName: qcModule, createdBy: "dashboard-quick" }),
      });
      const j = await r.json();
      if (!j.success) alert(j.error ?? "Komut gönderilemedi.");
    } finally {
      setQcSending(false);
    }
  }

  return (
    <div className="page-wrap page-wrap--wide">
      <PageHeader
        title="Dashboard"
        subtitle={`Operasyon özeti · ${apiEndpointCount} API endpoint · ${moduleCount} uzak modül`}
        live
      />

      <AlertBanner items={alerts} />

      <div className="stats-grid">
        <StatCard
          label="Çevrimiçi"
          value={onlineCount}
          hint={`${idleCount} boşta · ${offlineCount} kapalı`}
          tone="green"
          href="/heartbeats"
        />
        <StatCard
          label="Bekleyen komut"
          value={cmdPending}
          hint={cmdError > 0 ? `${cmdError} hatalı` : "Kuyruk temiz"}
          tone={cmdPending > 0 ? "yellow" : undefined}
          href="/komutlar"
        />
        <StatCard
          label="Aktif lisans"
          value={activeLicenseCount}
          hint={`${licenseCount} toplam kayıt`}
          tone="accent"
          href="/lisanslar"
        />
        <StatCard
          label="Uzak modül"
          value={activeModuleCount}
          hint={`${moduleCount} Neon DB`}
          tone="accent"
          href="/moduller"
        />
      </div>

      <div className="card" style={{ marginBottom: 16 }}>
        <div className="card-header">
          <span className="card-title">⚡ Hızlı Komut</span>
          <Link href="/komutlar" className="card-link">Komutlar</Link>
        </div>
        <div style={{ padding: "12px 18px", display: "flex", gap: 10, flexWrap: "wrap", alignItems: "flex-end" }}>
          <label style={{ fontSize: 12, display: "flex", flexDirection: "column", gap: 4 }}>
            MAC
            <select className="form-input" value={qcMac} onChange={(e) => setQcMac(e.target.value)} style={{ minWidth: 160 }}>
              {allMacs.map((m) => <option key={m} value={m}>{m}</option>)}
            </select>
          </label>
          <label style={{ fontSize: 12, display: "flex", flexDirection: "column", gap: 4, flex: 1, minWidth: 180 }}>
            Modül
            <select className="form-input" value={qcModule} onChange={(e) => setQcModule(e.target.value)}>
              {allModuleNames.map((m) => <option key={m} value={m}>{m}</option>)}
            </select>
          </label>
          <button type="button" className="btn btn-primary" onClick={sendQuickCommand} disabled={qcSending || !qcMac}>
            {qcSending ? "…" : "Gönder"}
          </button>
        </div>
      </div>

      {geoLocations.length > 0 && (
        <div className="card" style={{ marginBottom: 16 }}>
          <div className="card-header">
            <span className="card-title">🌍 Coğrafi IP Özeti</span>
            <Link href="/cihazlar" className="card-link">Cihazlar</Link>
          </div>
          <div style={{ padding: "8px 18px 14px", display: "flex", flexWrap: "wrap", gap: 8 }}>
            {geoLocations.map((g) => (
              <span key={g.mac} className="badge" style={{ fontSize: 11 }}>
                {g.hostname ?? g.mac} · {g.country}{g.city ? `, ${g.city}` : ""}
              </span>
            ))}
          </div>
        </div>
      )}

      <div className="quick-links">
        {[
          { href: "/komutlar", label: "Uzak Komutlar", icon: Terminal },
          { href: "/heartbeats", label: "Nabız", icon: Radio },
          { href: "/lisanslar", label: "Lisanslar", icon: KeyRound },
          { href: "/moduller", label: "Modüller", icon: Package },
        ].map(({ href, label, icon: Icon }) => (
          <Link key={href} href={href} className="quick-link card">
            <Icon size={18} />
            <span>{label}</span>
          </Link>
        ))}
      </div>

      <div className="grid-2">
        <div className="card">
          <div className="card-header">
            <span className="card-title">Son nabız</span>
            <Link href="/heartbeats" className="card-link">Tümü</Link>
          </div>
          <DataTable
            data={heartbeats.slice(0, 8)}
            pageSize={8}
            emptyTitle="Henüz heartbeat yok"
            rowKey={(h) => h.mac}
            columns={[
              {
                key: "st",
                header: "",
                width: "32px",
                render: (h) => <HeartbeatDot status={getHeartbeatStatus(h.last_seen)} />,
              },
              {
                key: "mac",
                header: "MAC",
                sortable: true,
                sortValue: (h) => h.mac,
                render: (h) => <span className="mono">{h.mac}</span>,
              },
              {
                key: "host",
                header: "PC",
                render: (h) => h.hostname ?? "—",
              },
              {
                key: "seen",
                header: "Son",
                sortable: true,
                sortValue: (h) => h.last_seen,
                render: (h) => <span className="text-muted">{timeAgo(h.last_seen)}</span>,
              },
            ]}
          />
        </div>

        <div className="card">
          <div className="card-header">
            <span className="card-title">Son komutlar</span>
            <Link href="/komutlar" className="card-link">Tümü</Link>
          </div>
          <DataTable
            data={commands.slice(0, 8)}
            pageSize={8}
            emptyTitle="Komut yok"
            rowKey={(c) => c.id}
            columns={[
              {
                key: "mod",
                header: "Modül",
                sortable: true,
                sortValue: (c) => c.moduleName,
                render: (c) => <span className="mono">{c.moduleName}</span>,
              },
              {
                key: "mac",
                header: "MAC",
                render: (c) => <span className="mono">{c.mac}</span>,
              },
              {
                key: "st",
                header: "Durum",
                render: (c) => <CommandStatusBadge status={c.status} />,
              },
            ]}
          />
        </div>
      </div>

      <div className="grid-2">
        <div className="card">
          <div className="card-header">
            <span className="card-title">En çok çalışan modüller</span>
            <Link href="/analitik" className="card-link">Analitik</Link>
          </div>
          {topModules.length === 0 ? (
            <p className="card-pad text-muted">Henüz çalışma verisi yok.</p>
          ) : (
            <div className="bar-list">
              {topModules.map((m) => {
                const max = topModules[0].runCount ?? 1;
                const pct = Math.round(((m.runCount ?? 0) / max) * 100);
                return (
                  <div key={m.methodName} className="bar-list-row">
                    <span className="mono bar-list-label">{m.methodName}</span>
                    <div className="bar-track">
                      <div className="bar-fill" style={{ width: `${pct}%` }} />
                    </div>
                    <span className="bar-list-value">{m.runCount}×</span>
                  </div>
                );
              })}
            </div>
          )}
        </div>

        <div className="card">
          <div className="card-header">
            <span className="card-title">Sistem özeti</span>
          </div>
          <dl className="summary-list">
            <div><dt>Cihaz snapshot</dt><dd>{snapshotCount}</dd></div>
            <div><dt>Modül çıktısı</dt><dd>{outputCount}</dd></div>
            <div><dt>Firma oto-modül</dt><dd>{firmAutoModuleCount}</dd></div>
            <div><dt>API endpoint</dt><dd>{apiEndpointCount}</dd></div>
          </dl>
          {globalChain && (
            <div className="card-pad card-pad--border">
              <div className="summary-chain-label">Global açılış zinciri</div>
              <p className="summary-chain">{globalChain}</p>
            </div>
          )}
        </div>
      </div>

      <div className="grid-2">
        <div className="card">
          <div className="card-header">
            <span className="card-title">Son lisanslar</span>
            <Link href="/lisanslar" className="card-link">Tümü</Link>
          </div>
          <DataTable
            data={licenses.slice(0, 6)}
            pageSize={6}
            emptyTitle="Lisans yok"
            rowKey={(l) => l.macAdresi}
            columns={[
              { key: "mac", header: "MAC", render: (l) => <span className="mono">{l.macAdresi}</span> },
              { key: "firma", header: "Firma", render: (l) => l.firmaAdi ?? "—" },
              { key: "lic", header: "Lisans", render: (l) => <LicenseBadge value={l.license} /> },
            ]}
          />
        </div>

        <div className="card">
          <div className="card-header">
            <span className="card-title">Son aktivite</span>
            <Link href="/loglar" className="card-link">Tümü</Link>
          </div>
          <DataTable
            data={logs}
            pageSize={8}
            emptyTitle="Log yok"
            rowKey={(l) => l.id}
            columns={[
              { key: "ev", header: "Olay", render: (l) => l.eventType },
              { key: "mac", header: "MAC", render: (l) => <span className="mono">{l.macAdresi ?? "—"}</span> },
              {
                key: "t",
                header: "Zaman",
                render: (l) => <span className="text-muted">{formatTR(l.createdAt)}</span>,
              },
            ]}
          />
        </div>
      </div>

      {cmdErrorRecent.length > 0 && (
        <div className="card">
          <div className="card-header">
            <span className="card-title">Son hatalı komutlar</span>
            <Link href="/komutlar?status=error" className="card-link">Filtrele</Link>
          </div>
          <DataTable
            data={cmdErrorRecent}
            pageSize={5}
            rowKey={(c) => c.id}
            columns={[
              { key: "mod", header: "Modül", render: (c) => <span className="mono">{c.moduleName}</span> },
              { key: "mac", header: "MAC", render: (c) => <span className="mono">{c.mac}</span> },
              { key: "err", header: "Hata", render: (c) => <span className="text-muted">{c.errorMsg ?? "—"}</span> },
            ]}
          />
        </div>
      )}
    </div>
  );
}
