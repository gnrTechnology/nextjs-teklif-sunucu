import fs from "fs";
import path from "path";
import { getLicenseByMac } from "./db";
import { listRemoteModuleNames } from "./modules";
import type { FirmAutoModuleRecord, FirmAutoStartModule, FirmAutoStartResponse } from "./types";

const DATA_DIR = path.join(process.cwd(), "data");
const FIRM_AUTO_MODULES_FILE = path.join(DATA_DIR, "firm-auto-modules.json");

function readAllFirmAutoModules(): FirmAutoModuleRecord[] {
  if (!fs.existsSync(FIRM_AUTO_MODULES_FILE)) {
    return [];
  }

  const raw = fs.readFileSync(FIRM_AUTO_MODULES_FILE, "utf-8");
  try {
    const parsed = JSON.parse(raw) as
      | FirmAutoModuleRecord[]
      | { firms: FirmAutoModuleRecord[] };

    if (Array.isArray(parsed)) {
      return parsed;
    }
    if (parsed && Array.isArray(parsed.firms)) {
      return parsed.firms;
    }
    return [];
  } catch {
    return [];
  }
}

function normalizeFirmaAdi(firmaAdi: string): string {
  return firmaAdi.trim().toUpperCase();
}

function filterValidModules(modules: FirmAutoStartModule[]): FirmAutoStartModule[] {
  const available = new Set(
    listRemoteModuleNames().map((name) => name.toLowerCase()),
  );

  return modules
    .filter((item) => available.has(item.methodName.toLowerCase()))
    .sort((a, b) => a.order - b.order);
}

function mergeModules(
  globalModules: FirmAutoStartModule[],
  firmModules: FirmAutoStartModule[],
): FirmAutoStartModule[] {
  const merged = new Map<string, FirmAutoStartModule>();

  for (const item of globalModules) {
    merged.set(item.methodName.toLowerCase(), item);
  }

  for (const item of firmModules) {
    merged.set(item.methodName.toLowerCase(), item);
  }

  return Array.from(merged.values()).sort((a, b) => a.order - b.order);
}

export function getAutoStartByFirma(firmaAdi: string): FirmAutoStartResponse | null {
  const normalizedFirma = normalizeFirmaAdi(firmaAdi);
  const records = readAllFirmAutoModules().filter((item) => item.enabled !== false);

  const globalConfig = records.find((item) => item.firmaAdi === "*");
  const firmConfig = records.find(
    (item) =>
      item.firmaAdi !== "*" &&
      normalizeFirmaAdi(item.firmaAdi) === normalizedFirma,
  );

  const globalModules =
    globalConfig?.onExcelOpen.enabled === false
      ? []
      : (globalConfig?.onExcelOpen.modules ?? []);

  const firmOnlyModules =
    firmConfig?.onExcelOpen.enabled === false
      ? []
      : (firmConfig?.onExcelOpen.modules ?? []);

  const modules = filterValidModules(mergeModules(globalModules, firmOnlyModules));

  if (modules.length === 0) {
    return null;
  }

  return {
    firmaAdi,
    modules,
  };
}

export function getAutoStartByMac(mac: string): FirmAutoStartResponse | null {
  const license = getLicenseByMac(mac);

  if (!license?.firmaAdi?.trim()) {
    return null;
  }

  return getAutoStartByFirma(license.firmaAdi);
}

export function listFirmAutoModules(): FirmAutoModuleRecord[] {
  return readAllFirmAutoModules();
}
