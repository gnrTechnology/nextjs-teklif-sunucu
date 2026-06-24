import { NextRequest } from "next/server";
import { errorResponse, jsonResponse } from "@/lib/api-response";
import { ensureModulesTable, listDbModules, listDbModulesPaginated, upsertDbModule, insertActivityLog } from "@/lib/db";
import type { ModuleUpsertBody } from "@/lib/types";

/** GET /api/modules — tüm modülleri listele (opsiyonel: ?search=&category=&offset=&limit=) */
export async function GET(request: NextRequest) {
  try {
    await ensureModulesTable();
    const sp = new URL(request.url).searchParams;
    const search = sp.get("search") ?? undefined;
    const category = sp.get("category") ?? undefined;
    const offset = sp.get("offset") ? Number(sp.get("offset")) : undefined;
    const limit = sp.get("limit") ? Number(sp.get("limit")) : undefined;

    if (search || category || offset !== undefined || limit !== undefined) {
      const result = await listDbModulesPaginated({ search, category, offset, limit });
      return jsonResponse({ success: true, ...result });
    }

    const modules = await listDbModules();
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
    await insertActivityLog({
      title: "Modül kaydedildi",
      detail: body.methodName.trim(),
      source: "dashboard/moduller",
    });
    return jsonResponse({ success: true, data: record }, 201);
  } catch (err) {
    console.error("[POST /api/modules]", err);
    return errorResponse("Modül kaydedilemedi.", 500);
  }
}
