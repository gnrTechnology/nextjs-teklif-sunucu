import path from "path";
import { jsonResponse, errorResponse } from "@/lib/api-response";
import { upsertFirmAutoModuleDb, ensureFirmAutoModulesTable, listFirmAutoModulesDb } from "@/lib/db";
import type { FirmAutoModuleRecord } from "@/lib/types";

/**
 * GET /api/firm-modules/seed
 * data/firm-auto-modules.json → Neon DB seed
 * Zaten kayıt varsa atlar (upsert — var olanı günceller).
 */
export async function GET() {
  try {
    const fs = await import("fs");
    const filePath = path.join(process.cwd(), "data", "firm-auto-modules.json");

    if (!fs.existsSync(filePath)) {
      return errorResponse("firm-auto-modules.json bulunamadı.", 404);
    }

    const raw = fs.readFileSync(filePath, "utf-8");
    const records: FirmAutoModuleRecord[] = JSON.parse(raw);

    if (!Array.isArray(records)) {
      return errorResponse("JSON formatı geçersiz (array bekleniyor).", 400);
    }

    await ensureFirmAutoModulesTable();
    let seeded = 0;
    for (const rec of records) {
      await upsertFirmAutoModuleDb(rec);
      seeded++;
    }

    const all = await listFirmAutoModulesDb();
    return jsonResponse({
      success: true,
      seeded,
      total: all.length,
      message: `${seeded} kayıt Neon DB'ye aktarıldı.`,
    });
  } catch (err) {
    return errorResponse(`Seed hatası: ${String(err)}`, 500);
  }
}
