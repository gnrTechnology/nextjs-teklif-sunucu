import {
  listFirmAutoModulesDb,
  getFirmAutoModuleDb,
  ensureFirmAutoModulesTable,
} from "./db";
import { getLicenseByMac } from "./db";
import { listRemoteModuleNames } from "./modules";
import type { FirmAutoModuleRecord, FirmAutoStartModule, FirmAutoStartResponse } from "./types";

let _ready = false;
async function ensureReady(): Promise<void> {
  if (_ready) return;
  await ensureFirmAutoModulesTable();
  _ready = true;
}

function normalizeFirmaAdi(firmaAdi: string): string {
  return firmaAdi.trim().toUpperCase();
}

function filterValidModules(
  modules: FirmAutoStartModule[],
  available: Set<string>,
): FirmAutoStartModule[] {
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

export async function getAutoStartByFirma(
  firmaAdi: string,
): Promise<FirmAutoStartResponse | null> {
  await ensureReady();

  const normalizedFirma = normalizeFirmaAdi(firmaAdi);
  const allRecords = (await listFirmAutoModulesDb()).filter((item) => item.enabled !== false);

  const globalConfig = allRecords.find((item) => item.firmaAdi === "*");
  const firmConfig = allRecords.find(
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

  const names = await listRemoteModuleNames();
  const available = new Set(names.map((n) => n.toLowerCase()));

  const modules = filterValidModules(mergeModules(globalModules, firmOnlyModules), available);

  if (modules.length === 0) {
    return null;
  }

  return { firmaAdi, modules };
}

export async function getAutoStartByMac(
  mac: string,
): Promise<FirmAutoStartResponse | null> {
  const license = await getLicenseByMac(mac);
  const firmaAdi = license?.firmaAdi?.trim() || "*";
  return getAutoStartByFirma(firmaAdi);
}

export async function listFirmAutoModules(): Promise<FirmAutoModuleRecord[]> {
  await ensureReady();
  return listFirmAutoModulesDb();
}

export async function getFirmAutoModule(firmaAdi: string): Promise<FirmAutoModuleRecord | undefined> {
  await ensureReady();
  return getFirmAutoModuleDb(firmaAdi);
}
