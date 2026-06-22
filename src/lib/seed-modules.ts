import fs from "fs";
import path from "path";
import { ensureModulesTable, upsertDbModule } from "@/lib/db";
import type { ModuleRecord } from "@/lib/types";

/**
 * modules.json → Neon DB upsert (tüm modüller).
 */
export async function seedModulesFromJson(): Promise<{
  seeded: number;
  errors: string[];
}> {
  const filePath = path.join(process.cwd(), "data", "modules.json");
  if (!fs.existsSync(filePath)) {
    throw new Error("data/modules.json bulunamadı.");
  }

  const raw = fs.readFileSync(filePath, "utf-8");
  const list = JSON.parse(raw) as ModuleRecord[];

  if (!Array.isArray(list)) {
    throw new Error("modules.json dizi formatında değil.");
  }

  await ensureModulesTable();

  let seeded = 0;
  const errors: string[] = [];

  for (const item of list) {
    if (!item.methodName || !item.code) continue;
    try {
      await upsertDbModule({
        methodName: item.methodName,
        description: item.description ?? "",
        category: item.category ?? "genel",
        code: item.code,
        active: item.active ?? true,
      });
      seeded++;
    } catch (e) {
      errors.push(`${item.methodName}: ${String(e)}`);
    }
  }

  return { seeded, errors };
}
