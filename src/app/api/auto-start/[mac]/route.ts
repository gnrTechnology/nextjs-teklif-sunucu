import { NextRequest } from "next/server";
import { errorResponse, jsonResponse } from "@/lib/api-response";
import { getAutoStartByMac } from "@/lib/firm-auto-modules";

type RouteContext = {
  params: Promise<{ mac: string }>;
};

/**
 * VBA: Excel açılışında MAC'e göre otomatik çalışacak modülleri döner.
 * GET /api/auto-start/{mac}/
 */
export async function GET(_request: NextRequest, context: RouteContext) {
  const { mac } = await context.params;
  const decodedMac = decodeURIComponent(mac);

  if (!decodedMac || decodedMac.length < 10) {
    return errorResponse("Geçersiz MAC adresi.", 400);
  }

  const autoStart = getAutoStartByMac(decodedMac);

  if (!autoStart) {
    return jsonResponse({
      success: true,
      data: {
        firmaAdi: null,
        modules: [],
      },
      message: "Bu MAC için firma kaydı veya otomatik modül tanımı bulunamadı.",
    });
  }

  return jsonResponse({
    success: true,
    data: autoStart,
  });
}
