/**
 * Modül erişim katmanı — Neon DB üzerinden çalışır.
 * modules.json yedek kaynak; eksik modüller otomatik senkronize edilir.
 */
import fs from "fs";
import path from "path";
import {
  getDbModuleByMethodName,
  listDbModules,
  upsertDbModule,
  ensureModulesTable,
} from "./db";
import type { ModuleRecord } from "./types";

const MODULES_JSON = path.join(process.cwd(), "data", "modules.json");

/** JSON'dan güncellenmesi gereken modüller (kod sık değişir) */
const ALWAYS_SYNC_FROM_JSON = new Set([
  "HeartbeatPing",
  "InstallTeklifAgent",
  "InstallCommandQueue",
  "GetCpuUsage",
  "GetNetworkSpeed",
]);

function readModulesJson(): ModuleRecord[] {
  if (!fs.existsSync(MODULES_JSON)) return [];
  try {
    const list = JSON.parse(fs.readFileSync(MODULES_JSON, "utf-8")) as ModuleRecord[];
    return Array.isArray(list) ? list : [];
  } catch {
    return [];
  }
}

/** Tek modülü modules.json'dan DB'ye upsert eder */
async function syncOneModuleFromJson(methodName: string): Promise<boolean> {
  const item = readModulesJson().find(
    (m) => m.methodName?.toLowerCase() === methodName.toLowerCase(),
  );
  if (!item?.methodName || !item.code) return false;

  await ensureModulesTable();
  await upsertDbModule({
    methodName: item.methodName,
    description: item.description ?? "",
    category: item.category ?? "genel",
    code: item.code,
    active: item.active ?? true,
  });
  return true;
}

/** Eksik modülleri JSON'dan toplu senkronize eder */
export async function syncMissingModulesFromJson(): Promise<number> {
  const list = readModulesJson();
  if (list.length === 0) return 0;

  await ensureModulesTable();
  const modules = await listDbModules();
  const existing = new Set(modules.map((m) => m.methodName.toLowerCase()));

  let synced = 0;
  for (const item of list) {
    if (!item.methodName || !item.code) continue;
    if (existing.has(item.methodName.toLowerCase())) continue;
    await upsertDbModule({
      methodName: item.methodName,
      description: item.description ?? "",
      category: item.category ?? "genel",
      code: item.code,
      active: item.active ?? true,
    });
    synced++;
  }
  return synced;
}

export async function getRemoteModuleCode(
  methodName: string,
): Promise<string | null> {
  const name = methodName.trim();
  let record = await getDbModuleByMethodName(name);

  const needsSync =
    !record || ALWAYS_SYNC_FROM_JSON.has(name);

  if (needsSync) {
    const ok = await syncOneModuleFromJson(name);
    if (ok) record = await getDbModuleByMethodName(name);
  }

  return record?.code ?? null;
}

export async function listModules(): Promise<ModuleRecord[]> {
  return listDbModules();
}

export async function listRemoteModuleNames(): Promise<string[]> {
  const modules = await listDbModules();
  return modules.filter((m) => m.active !== false).map((m) => m.methodName);
}
