import fs from "fs";
import path from "path";
import { NextRequest } from "next/server";
import { jsonResponse, errorResponse } from "@/lib/api-response";
import type { FirmAutoModuleRecord, FirmAutoStartModule } from "@/lib/types";

const FILE = path.join(process.cwd(), "data", "firm-auto-modules.json");

function readAll(): FirmAutoModuleRecord[] {
  if (!fs.existsSync(FILE)) return [];
  try {
    const raw = fs.readFileSync(FILE, "utf-8");
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : parsed.firms ?? [];
  } catch { return []; }
}
function writeAll(data: FirmAutoModuleRecord[]): void {
  fs.writeFileSync(FILE, JSON.stringify(data, null, 2), "utf-8");
}

type PatchBody = {
  description?: string;
  enabled?: boolean;
  onExcelOpenEnabled?: boolean;
  /** Modül ekle/güncelle */
  addModule?: { methodName: string; order?: number; delaySeconds?: number };
  /** Modül sil */
  removeModule?: string;
  /** Mevcut modülü güncelle */
  updateModule?: { methodName: string; order?: number; delaySeconds?: number };
  /** Tüm modülleri yeniden sırala */
  reorderModules?: { methodName: string; order: number }[];
};

/**
 * PATCH /api/firm-modules/[firmaAdi]
 * Firma ayarları + modül ekle/güncelle/çıkar/sırala
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

  const all = readAll();
  const idx = all.findIndex(
    (f) => f.firmaAdi.toLowerCase() === decoded.toLowerCase(),
  );
  if (idx === -1) return errorResponse("Firma bulunamadı.", 404);

  const firm = { ...all[idx] };

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

  // Modül çıkar
  if (body.removeModule) {
    const mods = (firm.onExcelOpen.modules ?? []).filter(
      (m) => m.methodName.toLowerCase() !== body.removeModule!.toLowerCase(),
    );
    // Sıraları yeniden ver
    mods.sort((a, b) => a.order - b.order).forEach((m, i) => (m.order = i + 1));
    firm.onExcelOpen = { ...firm.onExcelOpen, modules: mods };
  }

  // Toplu sıralama
  if (body.reorderModules) {
    const mods = firm.onExcelOpen.modules ?? [];
    body.reorderModules.forEach(({ methodName, order }) => {
      const m = mods.find((x) => x.methodName.toLowerCase() === methodName.toLowerCase());
      if (m) m.order = order;
    });
    mods.sort((a, b) => a.order - b.order);
    firm.onExcelOpen = { ...firm.onExcelOpen, modules: mods };
  }

  all[idx] = firm;
  writeAll(all);
  return jsonResponse({ success: true, data: firm });
}

/** DELETE /api/firm-modules/[firmaAdi] — firma kaydını sil */
export async function DELETE(
  _request: NextRequest,
  { params }: { params: Promise<{ firmaAdi: string }> },
) {
  const { firmaAdi } = await params;
  const decoded = decodeURIComponent(firmaAdi);

  if (decoded === "*") return errorResponse("Tüm-firma kaydı silinemez.", 403);

  const all = readAll();
  const filtered = all.filter(
    (f) => f.firmaAdi.toLowerCase() !== decoded.toLowerCase(),
  );
  if (filtered.length === all.length) return errorResponse("Firma bulunamadı.", 404);

  writeAll(filtered);
  return jsonResponse({ success: true });
}
