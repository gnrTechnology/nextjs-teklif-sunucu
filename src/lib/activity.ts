import type { ActivityCategory, UnifiedActivityItem } from "./types";
import {
  listLogs,
  listFolderWatchEvents,
  listHeartbeatLogs,
  listActivityLogs,
  listModuleOutputs,
  listClientCommands,
  listDeviceSnapshots,
} from "./db";

const LICENSE_CATEGORY: Record<string, Exclude<ActivityCategory, "all">> = {
  register: "lisans",
  update: "guncelleme",
  activate: "guncelleme",
  deactivate: "guncelleme",
  violation: "ihlal",
};

export const ACTIVITY_CATEGORY_LABELS: Record<
  Exclude<ActivityCategory, "all">,
  { label: string; icon: string; color: string }
> = {
  lisans: { label: "Lisans", icon: "🔑", color: "var(--accent)" },
  ihlal: { label: "İhlal", icon: "⚠", color: "var(--red)" },
  guncelleme: { label: "Güncelleme", icon: "🔄", color: "var(--yellow)" },
  dashboard: { label: "Dashboard", icon: "🖥", color: "var(--accent)" },
  modul: { label: "Modül Çıktısı", icon: "📤", color: "var(--green)" },
  heartbeat: { label: "Heartbeat", icon: "📡", color: "var(--green)" },
  komut: { label: "Uzak Komut", icon: "🎮", color: "var(--yellow)" },
  klasor: { label: "Klasör İzleme", icon: "📁", color: "var(--accent)" },
  cihaz: { label: "Cihaz Bilgisi", icon: "💻", color: "var(--text-muted)" },
};

export async function listUnifiedActivity(options?: {
  category?: ActivityCategory;
  limit?: number;
  mac?: string;
}): Promise<UnifiedActivityItem[]> {
  const limit = options?.limit ?? 400;
  const perSource = Math.ceil(limit / 4);

  const [
    licenseLogs,
    folderEvents,
    heartbeatLogs,
    dashboardLogs,
    moduleOutputs,
    commands,
    snapshots,
  ] = await Promise.all([
    listLogs(perSource),
    listFolderWatchEvents({ limit: perSource, mac: options?.mac }),
    listHeartbeatLogs({ limit: perSource, mac: options?.mac }),
    listActivityLogs({ limit: perSource }),
    listModuleOutputs({ limit: perSource, mac: options?.mac }),
    listClientCommands({ limit: perSource, mac: options?.mac }),
    listDeviceSnapshots(),
  ]);

  const items: UnifiedActivityItem[] = [];

  for (const log of licenseLogs) {
    const cat = LICENSE_CATEGORY[log.eventType] ?? "guncelleme";
    items.push({
      id: `license-${log.id}`,
      category: cat,
      title: eventTitle(log.eventType),
      detail: log.details,
      mac: log.macAdresi,
      source: log.eventType,
      createdAt: log.createdAt,
    });
  }

  for (const ev of folderEvents) {
    items.push({
      id: `folder-${ev.id}`,
      category: "klasor",
      title: folderEventTitle(ev.eventType, ev.fileName),
      detail: ev.detail ?? ev.filePath,
      mac: ev.mac,
      hostname: ev.hostname,
      source: ev.folderPath,
      createdAt: ev.createdAt,
    });
  }

  for (const hb of heartbeatLogs) {
    items.push({
      id: `hb-${hb.id}`,
      category: "heartbeat",
      title: "Heartbeat ping",
      detail: hb.userName ? `${hb.userName} · Excel v${hb.excelVersion ?? "?"}` : undefined,
      mac: hb.mac,
      hostname: hb.hostname,
      source: "HeartbeatPing",
      createdAt: hb.createdAt,
    });
  }

  for (const al of dashboardLogs) {
    items.push({
      id: `dash-${al.id}`,
      category: "dashboard",
      title: al.title,
      detail: al.detail,
      mac: al.mac,
      hostname: al.hostname,
      source: al.source,
      createdAt: al.createdAt,
    });
  }

  for (const out of moduleOutputs) {
    const preview = JSON.stringify(out.output).slice(0, 120);
    items.push({
      id: `modout-${out.id}`,
      category: "modul",
      title: `${out.moduleName} çıktısı`,
      detail: preview.length > 119 ? preview + "…" : preview,
      mac: out.mac,
      hostname: out.hostname,
      source: out.moduleName,
      createdAt: out.createdAt,
    });
  }

  for (const cmd of commands) {
    items.push({
      id: `cmd-${cmd.id}`,
      category: "komut",
      title: `${cmd.moduleName} — ${cmdStatusLabel(cmd.status)}`,
      detail: cmd.errorMsg ?? cmd.result ?? cmd.param ?? undefined,
      mac: cmd.mac,
      source: cmd.moduleName,
      createdAt: cmd.executedAt ?? cmd.createdAt,
    });
  }

  for (const snap of snapshots.slice(0, perSource)) {
    items.push({
      id: `dev-${snap.mac}`,
      category: "cihaz",
      title: "Cihaz bilgisi güncellendi",
      detail: snap.hostname ?? snap.mac,
      mac: snap.mac,
      hostname: snap.hostname,
      source: "CollectDeviceInfoServer",
      createdAt: snap.collectedAt,
    });
  }

  items.sort((a, b) => new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime());

  const filtered =
    options?.category && options.category !== "all"
      ? items.filter((i) => i.category === options.category)
      : items;

  return filtered.slice(0, limit);
}

function eventTitle(type: string): string {
  const map: Record<string, string> = {
    register: "Yeni lisans kaydı",
    update: "Lisans güncellendi",
    activate: "Lisans aktifleştirildi",
    deactivate: "Lisans pasifleştirildi",
    violation: "Lisans ihlali tespit edildi",
  };
  return map[type] ?? type;
}

function folderEventTitle(type: string, fileName?: string | null): string {
  const map: Record<string, string> = {
    created: "Yeni dosya",
    deleted: "Dosya silindi",
    modified: "Dosya değişti",
    started: "İzleme başlatıldı",
    scan: "Tarama tamamlandı",
  };
  const base = map[type] ?? type;
  return fileName ? `${base}: ${fileName}` : base;
}

function cmdStatusLabel(status: string): string {
  const map: Record<string, string> = {
    pending: "bekliyor",
    running: "çalışıyor",
    done: "tamamlandı",
    error: "hata",
  };
  return map[status] ?? status;
}

export function countByCategory(items: UnifiedActivityItem[]): Record<string, number> {
  const counts: Record<string, number> = {};
  for (const item of items) {
    counts[item.category] = (counts[item.category] ?? 0) + 1;
  }
  return counts;
}
