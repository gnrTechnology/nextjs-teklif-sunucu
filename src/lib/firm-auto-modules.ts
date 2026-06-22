import {
  listFirmAutoModulesDb,
  getFirmAutoModuleDb,
  ensureFirmAutoModulesTable,
  upsertFirmAutoModuleDb,
} from "./db";
import { getLicenseByMac } from "./db";
import { listRemoteModuleNames } from "./modules";
import type { FirmAutoModuleRecord, FirmAutoStartModule, FirmAutoStartResponse } from "./types";

/** İlk kez çağrıldığında JSON'dan Neon'a seed eder (idempotent) */
async function seedFromJsonIfEmpty(): Promise<void> {
  try {
    // Sadece Node.js ortamında (sunucu) çalışır
    const { default: fs } = await import("fs");
    const { default: path } = await import("path");
    const filePath = path.join(process.cwd(), "data", "firm-auto-modules.json");
    if (!fs.existsSync(filePath)) return;
    const raw = fs.readFileSync(filePath, "utf-8");
    const records: FirmAutoModuleRecord[] = JSON.parse(raw);
    if (!Array.isArray(records)) return;
    for (const rec of records) {
      await upsertFirmAutoModuleDb(rec);
    }
  } catch {
    /* seed başarısız olursa sessizce geç */
  }
}

let _seeded = false;
async function ensureSeeded(): Promise<void> {
  if (_seeded) return;
  await ensureFirmAutoModulesTable();
  const existing = await listFirmAutoModulesDb();
  if (existing.length === 0) {
    await seedFromJsonIfEmpty();
  }
  _seeded = true;
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
  await ensureSeeded();

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
  await ensureSeeded();
  return listFirmAutoModulesDb();
}

export async function getFirmAutoModule(firmaAdi: string): Promise<FirmAutoModuleRecord | undefined> {
  await ensureSeeded();
  return getFirmAutoModuleDb(firmaAdi);
}
