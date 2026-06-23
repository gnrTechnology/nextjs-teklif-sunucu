import type { ClientCommand } from "./db";
import type { FolderWatchHealth } from "./types";
import { elapsedSecSince } from "./date-utils";

export type StuckState = "ok" | "slow" | "stuck" | "alive_bg";

/** Modülün normal bitmesi beklenen süre (sn) — aşılırsa takılı uyarısı */
const MODULE_EXPECTED_SEC: Record<string, number> = {
  WatchFolderServer: 30,
  InstallCommandQueue: 20,
  HeartbeatPing: 15,
  GetCpuInfo: 45,
  InstallTeklifAgent: 30,
};

const DEFAULT_EXPECTED_SEC = 60;

export type CommandProgressView = {
  pct: number;
  label: string;
  stuckState: StuckState;
};

export function getCommandProgressView(
  cmd: ClientCommand,
  health?: FolderWatchHealth | null,
): CommandProgressView {
  if (cmd.status === "pending") {
    return { pct: 0, label: "Sırada bekliyor", stuckState: "ok" };
  }
  if (cmd.status === "done") {
    return {
      pct: 100,
      label: cmd.progressLabel ?? "Tamamlandı",
      stuckState: "ok",
    };
  }
  if (cmd.status === "error") {
    return {
      pct: cmd.progressPct ?? 0,
      label: cmd.errorMsg ?? cmd.progressLabel ?? "Hata",
      stuckState: "stuck",
    };
  }

  let pct = cmd.progressPct ?? 15;
  let label = cmd.progressLabel ?? "Excel'de çalıştırılıyor";
  let stuckState: StuckState = "ok";

  const runSec = elapsedSecSince(cmd.executedAt) ?? 0;
  const progStaleSec = elapsedSecSince(cmd.progressAt);
  const expected = MODULE_EXPECTED_SEC[cmd.moduleName] ?? DEFAULT_EXPECTED_SEC;

  if (cmd.moduleName === "WatchFolderServer") {
    if (health?.isAlive) {
      return {
        pct: 100,
        label: "Arka planda izleme aktif",
        stuckState: "alive_bg",
      };
    }
    if (health?.lastEventType === "started") {
      pct = Math.max(pct, 75);
      label = "İzleme başlatıldı — canlı ping bekleniyor";
    } else if (runSec > 20 && pct < 40) {
      pct = Math.max(pct, 25);
      label = "Modül çalışıyor — sunucudan sinyal bekleniyor";
    }
  }

  if (runSec > expected * 2 && pct < 85 && stuckState !== "alive_bg") {
    stuckState = "stuck";
    if (cmd.moduleName !== "WatchFolderServer" || !health?.lastPingAt) {
      label = `${label} · yanıt yok (${runSec} sn)`;
    }
  } else if (runSec > expected && pct < 60) {
    stuckState = "slow";
  }

  if (progStaleSec != null && progStaleSec > 90 && pct < 95) {
    if (stuckState === "ok") stuckState = "slow";
  }

  return { pct: Math.min(100, Math.max(0, pct)), label, stuckState };
}

export function stuckStateColor(state: StuckState): string {
  switch (state) {
    case "stuck":
      return "#ef4444";
    case "slow":
      return "#f59e0b";
    case "alive_bg":
      return "#10b981";
    default:
      return "#3b82f6";
  }
}

export function stuckStateHint(state: StuckState): string | null {
  switch (state) {
    case "stuck":
      return "Yanıt yok — modül takılı olabilir. Excel açık mı, break modunda mı kontrol edin.";
    case "slow":
      return "Beklenenden uzun sürüyor…";
    case "alive_bg":
      return "Komut satırı takılı kalsa bile arka plan servisi çalışıyor.";
    default:
      return null;
  }
}
