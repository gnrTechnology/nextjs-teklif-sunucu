import { NextRequest } from "next/server";
import { errorResponse, jsonResponse } from "@/lib/api-response";
import { getLicenseByMac, normalizeMac } from "@/lib/db";

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
      return jsonResponse(
        { success: false, message: "Bu MAC adresi için lisans bulunamadı.", data: null },
        404,
      );
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
