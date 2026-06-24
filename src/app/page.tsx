export const dynamic = "force-dynamic";

import Link from "next/link";
import {
  listLicenses,
  listLogs,
  listHeartbeats,
  listDeviceSnapshots,
  listClientCommands,
  listModuleOutputs,
  listDbModules,
} from "@/lib/db";
import { formatTR, getHeartbeatStatus, timeAgo } from "@/lib/date-utils";
import { listFirmAutoModules } from "@/lib/firm-auto-modules";
import { getApiCatalogStats } from "@/lib/api-catalog";
import PageHeader from "./components/ui/PageHeader";
import StatCard from "./components/ui/StatCard";
import AlertBanner, { type AlertItem } from "./components/ui/AlertBanner";
import DataTable from "./components/ui/DataTable";
import {
  CommandStatusBadge,
  HeartbeatDot,
  isLicenseActive,
  LicenseBadge,
} from "@/lib/status-badges";
import { Radio, Terminal, KeyRound, Package } from "lucide-react";

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
  ] = await Promise.all([
    listLicenses(),
    listDbModules(),
    listFirmAutoModules(),
    listLogs(8),
    listHeartbeats(),
    listDeviceSnapshots(),
    listClientCommands({ limit: 200 }),
    listModuleOutputs({ limit: 50 }),
  ]);

  const apiStats = getApiCatalogStats();
  const activeLicenses = licenses.filter((l) => isLicenseActive(l.license));
  const inactiveLicenses = licenses.length - activeLicenses.length;

  const onlineCount = heartbeats.filter((h) => getHeartbeatStatus(h.last_seen) === "online").length;
  const idleCount = heartbeats.filter((h) => getHeartbeatStatus(h.last_seen) === "idle").length;
  const offlineCount = heartbeats.length - onlineCount - idleCount;

  const cmdPending = commands.filter((c) => c.status === "pending" || c.status === "running").length;
  const cmdError = commands.filter((c) => c.status === "error").length;
  const cmdErrorRecent = commands.filter((c) => c.status === "error").slice(0, 5);

  const topModules = [...modules]
    .filter((m) => (m.runCount ?? 0) > 0)
    .sort((a, b) => (b.runCount ?? 0) - (a.runCount ?? 0))
    .slice(0, 5);

  const alerts: AlertItem[] = [];
  if (offlineCount > 0) {
    alerts.push({
      id: "offline",
      tone: offlineCount > 3 ? "danger" : "warning",
      message: `${offlineCount} cihaz çevrimdışı (>1 saat nabız yok)`,
      href: "/heartbeats",
    });
  }
  if (cmdError > 0) {
    alerts.push({
      id: "cmd-err",
      tone: "danger",
      message: `${cmdError} komut hata durumunda`,
      href: "/komutlar?status=error",
    });
  }
  if (cmdPending > 5) {
    alerts.push({
      id: "cmd-pending",
      tone: "warning",
      message: `${cmdPending} komut kuyrukta bekliyor`,
      href: "/komutlar",
    });
  }
  if (inactiveLicenses > 0) {
    alerts.push({
      id: "license",
      tone: "warning",
      message: `${inactiveLicenses} pasif lisans kaydı`,
      href: "/lisanslar",
    });
  }

  return (
    <div className="page-wrap page-wrap--wide">
      <PageHeader
        title="Dashboard"
        subtitle={`Operasyon özeti · ${apiStats.total} API endpoint · ${modules.length} uzak modül`}
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
          value={activeLicenses.length}
          hint={`${licenses.length} toplam kayıt`}
          tone="accent"
          href="/lisanslar"
        />
        <StatCard
          label="Uzak modül"
          value={modules.filter((m) => m.active !== false).length}
          hint={`${modules.length} Neon DB`}
          tone="accent"
          href="/moduller"
        />
      </div>

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
            <div><dt>Cihaz snapshot</dt><dd>{snapshots.length}</dd></div>
            <div><dt>Modül çıktısı</dt><dd>{outputs.length}</dd></div>
            <div><dt>Firma oto-modül</dt><dd>{firmAutoModules.length}</dd></div>
            <div><dt>API endpoint</dt><dd>{apiStats.total}</dd></div>
          </dl>
          {firmAutoModules.length > 0 && (
            <div className="card-pad card-pad--border">
              <div className="summary-chain-label">Global açılış zinciri</div>
              <p className="summary-chain">
                {(firmAutoModules.find((f) => f.firmaAdi === "*")?.onExcelOpen.modules ?? [])
                  .sort((a, b) => a.order - b.order)
                  .map((m) => m.methodName + (m.runOnce ? " (1×)" : ""))
                  .join(" → ") || "—"}
              </p>
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
