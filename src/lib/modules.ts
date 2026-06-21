import fs from "fs";
import path from "path";
import type { ModuleRecord } from "./types";

const DATA_DIR = path.join(process.cwd(), "data");
const MODULES_FILE = path.join(DATA_DIR, "modules.json");

function ensureDataDir(): void {
  if (!fs.existsSync(DATA_DIR)) {
    fs.mkdirSync(DATA_DIR, { recursive: true });
  }
}

function readAllModules(): ModuleRecord[] {
  ensureDataDir();
  if (!fs.existsSync(MODULES_FILE)) {
    return [];
  }

  const raw = fs.readFileSync(MODULES_FILE, "utf-8");
  try {
    const parsed = JSON.parse(raw) as ModuleRecord[] | { modules: ModuleRecord[] };
    if (Array.isArray(parsed)) {
      return parsed;
    }
    if (parsed && Array.isArray(parsed.modules)) {
      return parsed.modules;
    }
    return [];
  } catch {
    return [];
  }
}

export function getRemoteModuleCode(methodName: string): string | null {
  const normalized = methodName.trim();
  const records = readAllModules().filter(
    (item) => item.active !== false && item.code.trim() !== "",
  );

  const record =
    records.find((item) => item.methodName === normalized) ??
    records.find(
      (item) => item.methodName.toLowerCase() === normalized.toLowerCase(),
    );

  return record?.code ?? null;
}

export function listModules(): ModuleRecord[] {
  return readAllModules().filter((item) => item.active !== false);
}

export function listRemoteModuleNames(): string[] {
  return listModules().map((item) => item.methodName);
}
