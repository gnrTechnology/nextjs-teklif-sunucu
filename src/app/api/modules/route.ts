import { NextRequest } from "next/server";
import { errorResponse, jsonResponse } from "@/lib/api-response";
import { listDbModules, upsertDbModule, ensureModulesTable } from "@/lib/db";
import type { ModuleUpsertBody } from "@/lib/types";

/** GET /api/modules — tüm modülleri listele */
export async function GET() {
  try {
    await ensureModulesTable();
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
    return jsonResponse({ success: true, data: record }, 201);
  } catch (err) {
    console.error("[POST /api/modules]", err);
    return errorResponse("Modül kaydedilemedi.", 500);
  }
}
