import fs from "fs";
import path from "path";
import { NextRequest } from "next/server";
import { errorResponse, jsonResponse } from "@/lib/api-response";
import { listDbModules, upsertDbModule, ensureModulesTable } from "@/lib/db";
import type { ModuleUpsertBody, ModuleRecord } from "@/lib/types";

/** GET /api/modules — tüm modülleri listele; tablo yoksa oluştur + JSON'dan seed et */
export async function GET() {
  try {
    await ensureModulesTable();
    let modules = await listDbModules();

    // modules.json'daki eksik modülleri DB'ye ekle (INSERT ON CONFLICT DO NOTHING)
    const filePath = path.join(process.cwd(), "data", "modules.json");
    if (fs.existsSync(filePath)) {
      const list = JSON.parse(fs.readFileSync(filePath, "utf-8")) as ModuleRecord[];
      const existing = new Set(modules.map((m) => m.methodName.toLowerCase()));
      const newItems = list.filter(
        (item) => item.methodName && item.code && !existing.has(item.methodName.toLowerCase()),
      );
      if (newItems.length > 0) {
        for (const item of newItems) {
          await upsertDbModule({
            methodName: item.methodName,
            description: item.description ?? "",
            category: item.category ?? "genel",
            code: item.code,
            active: item.active ?? true,
          });
        }
        modules = await listDbModules();
      }
    }

    return jsonResponse({ success: true, data: modules });
  } catch (err) {
    console.error("[GET /api/modules]", err);
    return errorResponse("Modüller alınamadı.", 500);
  }
}

/** POST /api/modules — yeni modül oluştur veya güncelle (upsert) */
export async function POST(request: NextRequest) {
  let body: ModuleUpsertBody;
  try {
    body = (await request.json()) as ModuleUpsertBody;
  } catch {
    return errorResponse("Geçersiz JSON.", 400);
  }

  if (!body.methodName?.trim()) {
    return errorResponse("methodName zorunludur.", 400);
  }
  if (!body.code?.trim()) {
    return errorResponse("code zorunludur.", 400);
  }

  try {
    await ensureModulesTable();
    const record = await upsertDbModule(body);
    return jsonResponse({ success: true, data: record }, 201);
  } catch (err) {
    console.error("[POST /api/modules]", err);
    return errorResponse("Modül kaydedilemedi.", 500);
  }
}
