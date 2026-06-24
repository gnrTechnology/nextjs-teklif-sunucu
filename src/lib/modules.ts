/**
 * Modül erişim katmanı — Neon DB üzerinden çalışır.
 * Kaynak kod: data/modules-source/*.bas (Neon MCP ile DB'ye yazılır).
 */
import { getDbModuleByMethodName, listDbModules } from "./db";
import type { ModuleRecord } from "./types";

export async function getRemoteModuleCode(
  methodName: string,
): Promise<string | null> {
  const record = await getDbModuleByMethodName(methodName.trim());
  return record?.code ?? null;
}

export async function listModules(): Promise<ModuleRecord[]> {
  return listDbModules();
}

export async function listRemoteModuleNames(): Promise<string[]> {
  const modules = await listDbModules();
  return modules.filter((m) => m.active !== false).map((m) => m.methodName);
}
