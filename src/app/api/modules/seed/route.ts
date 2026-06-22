import fs from "fs";
import path from "path";
import { jsonResponse, errorResponse } from "@/lib/api-response";
import { ensureModulesTable, upsertDbModule } from "@/lib/db";
import type { ModuleRecord } from "@/lib/types";

/**
 * POST /api/modules/seed
 * modules.json içindeki tüm kayıtları Neon DB'ye aktarır.
 * Tek seferlik çalıştırılır; mevcut kayıtları günceller (upsert).
 */
export async function POST() {
  try {
    const filePath = path.join(process.cwd(), "data", "modules.json");
    if (!fs.existsSync(filePath)) {
      return errorResponse("data/modules.json bulunamadı.", 404);
    }

    const raw = fs.readFileSync(filePath, "utf-8");
    const list = JSON.parse(raw) as ModuleRecord[];

    if (!Array.isArray(list)) {
      return errorResponse("modules.json dizi formatında değil.", 400);
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

    return jsonResponse({
      success: true,
      seeded,
      errors: errors.length > 0 ? errors : undefined,
      message: `${seeded} modül Neon DB'ye aktarıldı.`,
    });
  } catch (err) {
    console.error("[POST /api/modules/seed]", err);
    return errorResponse("Seed başarısız.", 500);
  }
}
