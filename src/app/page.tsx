export const dynamic = "force-dynamic";

import {
  listLicenses,
  listLogs,
  listHeartbeats,
  listDeviceSnapshots,
  listClientCommands,
  listModuleOutputs,
  listDbModules,
} from "@/lib/db";
import { getHeartbeatStatus } from "@/lib/date-utils";
import { listFirmAutoModules } from "@/lib/firm-auto-modules";
import { getApiCatalogStats } from "@/lib/api-catalog";
import { isLicenseActive } from "@/lib/status-badges";
import { lookupGeo } from "@/lib/geoip";
import type { AlertItem } from "./components/ui/AlertBanner";
import DashboardClient from "./components/DashboardClient";

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
    .slice(0, 5)
    .map((m) => ({ methodName: m.methodName, runCount: m.runCount }));

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

  const globalChain =
    (firmAutoModules.find((f) => f.firmaAdi === "*")?.onExcelOpen.modules ?? [])
      .sort((a, b) => a.order - b.order)
      .map((m) => m.methodName + (m.runOnce ? " (1×)" : ""))
      .join(" → ") || "";

  const allMacs = heartbeats.map((h) => h.mac);
  const allModuleNames = modules.filter((m) => m.active !== false).map((m) => m.methodName);

  const geoLocations = (
    await Promise.all(
      snapshots
        .filter((s) => s.data.publicIp)
        .slice(0, 15)
        .map(async (s) => {
          const ip = String(s.data.publicIp);
          const geo = await lookupGeo(ip);
          if (!geo) return null;
          return {
            mac: s.mac,
            hostname: s.hostname,
            ip,
            country: geo.country,
            city: geo.city,
          };
        }),
    )
  ).filter((g): g is NonNullable<typeof g> => g != null);

  return (
    <DashboardClient
      apiEndpointCount={apiStats.total}
      moduleCount={modules.length}
      activeModuleCount={modules.filter((m) => m.active !== false).length}
      onlineCount={onlineCount}
      idleCount={idleCount}
      offlineCount={offlineCount}
      cmdPending={cmdPending}
      cmdError={cmdError}
      activeLicenseCount={activeLicenses.length}
      licenseCount={licenses.length}
      alerts={alerts}
      heartbeats={heartbeats}
      commands={commands}
      licenses={licenses}
      logs={logs}
      topModules={topModules}
      snapshotCount={snapshots.length}
      outputCount={outputs.length}
      firmAutoModuleCount={firmAutoModules.length}
      globalChain={globalChain}
      cmdErrorRecent={cmdErrorRecent}
      allMacs={allMacs}
      allModuleNames={allModuleNames}
      geoLocations={geoLocations}
    />
  );
}
