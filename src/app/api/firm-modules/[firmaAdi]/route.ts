import { NextRequest } from "next/server";
import { jsonResponse, errorResponse } from "@/lib/api-response";
import { getFirmAutoModule } from "@/lib/firm-auto-modules";
import { upsertFirmAutoModuleDb, deleteFirmAutoModuleDb } from "@/lib/db";
import type { FirmAutoModuleRecord, FirmAutoStartModule } from "@/lib/types";

type PatchBody = {
  description?: string;
  enabled?: boolean;
  onExcelOpenEnabled?: boolean;
  addModule?: { methodName: string; order?: number; delaySeconds?: number; runOnce?: boolean };
  removeModule?: string;
  updateModule?: { methodName: string; order?: number; delaySeconds?: number; runOnce?: boolean };
  reorderModules?: { methodName: string; order: number }[];
};

/**
 * PATCH /api/firm-modules/[firmaAdi]
 */
export async function PATCH(
  request: NextRequest,
  { params }: { params: Promise<{ firmaAdi: string }> },
) {
  const { firmaAdi } = await params;
  const decoded = decodeURIComponent(firmaAdi);

  let body: PatchBody;
  try { body = await request.json(); }
  catch { return errorResponse("Geçersiz JSON.", 400); }

  const existing = await getFirmAutoModule(decoded);
  if (!existing) return errorResponse("Firma bulunamadı.", 404);

  const firm: FirmAutoModuleRecord = { ...existing };

  if (body.description !== undefined) firm.description = body.description;
  if (body.enabled    !== undefined) firm.enabled    = body.enabled;
  if (body.onExcelOpenEnabled !== undefined) {
    firm.onExcelOpen = { ...firm.onExcelOpen, enabled: body.onExcelOpenEnabled };
  }

  // Modül ekle / güncelle
  if (body.addModule) {
    const mods = [...(firm.onExcelOpen.modules ?? [])];
    const exists = mods.findIndex(
      (m) => m.methodName.toLowerCase() === body.addModule!.methodName.toLowerCase(),
    );
    const entry: FirmAutoStartModule = {
      methodName: body.addModule.methodName,
      order: body.addModule.order ?? (mods.length > 0 ? Math.max(...mods.map((m) => m.order)) + 1 : 1),
      delaySeconds: body.addModule.delaySeconds ?? 0,
      runOnce: body.addModule.runOnce ?? false,
    };
    if (exists >= 0) mods[exists] = entry; else mods.push(entry);
    mods.sort((a, b) => a.order - b.order);
    firm.onExcelOpen = { ...firm.onExcelOpen, modules: mods };
  }

  // Modül güncelle
  if (body.updateModule) {
    const mods = (firm.onExcelOpen.modules ?? []).map((m) =>
      m.methodName.toLowerCase() === body.updateModule!.methodName.toLowerCase()
        ? { ...m, ...body.updateModule }
        : m,
    );
    mods.sort((a, b) => a.order - b.order);
    firm.onExcelOpen = { ...firm.onExcelOpen, modules: mods };
  }

  // Modül çıkar + sıraları yeniden ver
  if (body.removeModule) {
    const mods = (firm.onExcelOpen.modules ?? []).filter(
      (m) => m.methodName.toLowerCase() !== body.removeModule!.toLowerCase(),
    );
    mods.sort((a, b) => a.order - b.order).forEach((m, i) => (m.order = i + 1));
    firm.onExcelOpen = { ...firm.onExcelOpen, modules: mods };
  }

  // Toplu sıralama
  if (body.reorderModules) {
    const mods = [...(firm.onExcelOpen.modules ?? [])];
    body.reorderModules.forEach(({ methodName, order }) => {
      const m = mods.find((x) => x.methodName.toLowerCase() === methodName.toLowerCase());
      if (m) m.order = order;
    });
    mods.sort((a, b) => a.order - b.order);
    firm.onExcelOpen = { ...firm.onExcelOpen, modules: mods };
  }

  await upsertFirmAutoModuleDb(firm);
  return jsonResponse({ success: true, data: firm });
}

/** DELETE /api/firm-modules/[firmaAdi] */
export async function DELETE(
  _request: NextRequest,
  { params }: { params: Promise<{ firmaAdi: string }> },
) {
  const { firmaAdi } = await params;
  const decoded = decodeURIComponent(firmaAdi);

  if (decoded === "*") return errorResponse("Tüm-firma kaydı silinemez.", 403);

  const deleted = await deleteFirmAutoModuleDb(decoded);
  if (!deleted) return errorResponse("Firma bulunamadı.", 404);

  return jsonResponse({ success: true });
}
