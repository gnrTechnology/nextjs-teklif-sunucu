import { NextRequest } from "next/server";
import { errorResponse, jsonResponse } from "@/lib/api-response";
import { getLicenseByMac, insertLog, normalizeMac, toggleLicense } from "@/lib/db";

type RouteContext = {
  params: Promise<{ mac: string }>;
};

/**
 * VBA: GetLicenseStatus()
 * GET http://host:3000/api/license/{macAddress}
 */
export async function GET(_request: NextRequest, context: RouteContext) {
  const { mac } = await context.params;
  const decodedMac = decodeURIComponent(mac);

  if (!decodedMac || decodedMac.length < 10) {
    return errorResponse("Geçersiz MAC adresi.", 400);
  }

  try {
    const record = await getLicenseByMac(decodedMac);

    if (!record) {
      return jsonResponse({
        success: false,
        message: "Bu MAC adresi için lisans bulunamadı.",
        data: null,
      });
    }

    return jsonResponse({
      success: true,
      data: {
        macAdresi: normalizeMac(record.macAdresi),
        license: record.license,
        firmaAdi: record.firmaAdi ?? null,
        userAdi: record.userAdi ?? null,
        dosyaAdi: record.dosyaAdi ?? null,
        ipAdresi: record.ipAdresi ?? null,
        updatedAt: record.updatedAt,
      },
    });
  } catch (err) {
    console.error("[GET /api/license/mac] DB hatası:", err);
    return errorResponse("Sunucu hatası.", 500, {
      detail: err instanceof Error ? err.message : String(err),
    });
  }
}

/**
 * Dashboard: Lisans aktif/pasif toggle
 * PATCH /api/license/{mac}  body: { license: "true" | "false" }
 */
export async function PATCH(request: NextRequest, context: RouteContext) {
  const { mac } = await context.params;
  const decodedMac = decodeURIComponent(mac);

  let body: { license: "true" | "false" };
  try {
    body = await request.json();
  } catch {
    return errorResponse("Geçersiz JSON.", 400);
  }

  if (body.license !== "true" && body.license !== "false") {
    return errorResponse('license alanı "true" veya "false" olmalı.', 400);
  }

  try {
    const record = await toggleLicense(decodedMac, body.license);
    if (!record) return errorResponse("Kayıt bulunamadı.", 404);

    await insertLog({
      macAdresi: record.macAdresi,
      firmaAdi: record.firmaAdi,
      userAdi: record.userAdi,
      eventType: body.license === "true" ? "activate" : "deactivate",
      details: `Dashboard üzerinden ${body.license === "true" ? "aktifleştirildi" : "pasifleştirildi"}.`,
    });

    return jsonResponse({ success: true, data: { license: record.license } });
  } catch (err) {
    console.error("[PATCH /api/license/mac] DB hatası:", err);
    return errorResponse("Sunucu hatası.", 500, {
      detail: err instanceof Error ? err.message : String(err),
    });
  }
}
