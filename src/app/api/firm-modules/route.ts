import { NextRequest } from "next/server";
import { jsonResponse, errorResponse } from "@/lib/api-response";
import {
  listFirmAutoModules,
  getFirmAutoModule,
} from "@/lib/firm-auto-modules";
import { upsertFirmAutoModuleDb } from "@/lib/db";
import type { FirmAutoModuleRecord } from "@/lib/types";

/** GET /api/firm-modules */
export async function GET() {
  const data = await listFirmAutoModules();
  return jsonResponse({ success: true, data });
}

/** POST /api/firm-modules — yeni firma ekle */
export async function POST(request: NextRequest) {
  let body: { firmaAdi: string; description?: string; enabled?: boolean };
  try {
    body = await request.json();
  } catch {
    return errorResponse("Geçersiz JSON.", 400);
  }

  if (!body.firmaAdi?.trim()) return errorResponse("firmaAdi zorunludur.", 400);

  const existing = await getFirmAutoModule(body.firmaAdi.trim());
  if (existing) return errorResponse("Bu firma zaten mevcut.", 409);

  const newItem: FirmAutoModuleRecord = {
    firmaAdi: body.firmaAdi.trim(),
    description: body.description ?? "",
    enabled: body.enabled ?? true,
    onExcelOpen: { enabled: true, modules: [] },
  };

  await upsertFirmAutoModuleDb(newItem);
  return jsonResponse({ success: true, data: newItem }, 201);
}
