import { NextRequest } from "next/server";
import { errorResponse, jsonResponse } from "@/lib/api-response";
import { updateDbModule, deleteDbModule } from "@/lib/db";
import type { ModuleUpsertBody } from "@/lib/types";

/** PATCH /api/modules/[id] — modül alanlarını güncelle */
export async function PATCH(
  request: NextRequest,
  { params }: { params: Promise<{ id: string }> },
) {
  const { id } = await params;
  const numId = Number(id);
  if (isNaN(numId)) return errorResponse("Geçersiz id.", 400);

  let body: Partial<ModuleUpsertBody>;
  try {
    body = (await request.json()) as Partial<ModuleUpsertBody>;
  } catch {
    return errorResponse("Geçersiz JSON.", 400);
  }

  try {
    const record = await updateDbModule(numId, body);
    if (!record) return errorResponse("Modül bulunamadı.", 404);
    return jsonResponse({ success: true, data: record });
  } catch (err) {
    console.error("[PATCH /api/modules]", err);
    return errorResponse("Güncellenemedi.", 500);
  }
}

/** DELETE /api/modules/[id] — modülü sil */
export async function DELETE(
  _request: NextRequest,
  { params }: { params: Promise<{ id: string }> },
) {
  const { id } = await params;
  const numId = Number(id);
  if (isNaN(numId)) return errorResponse("Geçersiz id.", 400);

  try {
    const ok = await deleteDbModule(numId);
    if (!ok) return errorResponse("Modül bulunamadı.", 404);
    return jsonResponse({ success: true });
  } catch (err) {
    console.error("[DELETE /api/modules]", err);
    return errorResponse("Silinemedi.", 500);
  }
}
