import { NextRequest } from "next/server";
import { jsonResponse, errorResponse } from "@/lib/api-response";
import { upsertDeviceSnapshot, listDeviceSnapshots, getDeviceSnapshot } from "@/lib/db";

/**
 * POST /api/device-info
 * VBA CollectDeviceInfoServer modülü buraya POST atar.
 * Body: { mac, hostname?, firmaAdi?, data: { key: value, ... } }
 */
export async function POST(request: NextRequest) {
  let body: {
    mac: string;
    hostname?: string;
    firmaAdi?: string;
    data: Record<string, unknown>;
  };

  try {
    body = await request.json();
  } catch {
    return errorResponse("Geçersiz JSON.", 400);
  }

  if (!body.mac?.trim()) return errorResponse("mac alanı zorunludur.", 400);
  if (!body.data || typeof body.data !== "object") return errorResponse("data objesi zorunludur.", 400);

  /* IP adresini header'dan otomatik ekle */
  const ip =
    request.headers.get("x-forwarded-for")?.split(",")[0]?.trim() ??
    request.headers.get("x-real-ip") ??
    null;

  await upsertDeviceSnapshot({
    mac: body.mac.trim(),
    hostname: body.hostname ?? null,
    firmaAdi: body.firmaAdi ?? null,
    data: { ...body.data, _ip: ip, _collectedAt: new Date().toISOString() },
  });

  return jsonResponse({ success: true, message: "Cihaz verisi kaydedildi." });
}

/**
 * GET /api/device-info          → tüm cihazlar
 * GET /api/device-info?mac=...  → tek cihaz
 */
export async function GET(request: NextRequest) {
  const mac = new URL(request.url).searchParams.get("mac");

  if (mac) {
    const snap = await getDeviceSnapshot(mac);
    if (!snap) return errorResponse("Cihaz bulunamadı.", 404);
    return jsonResponse({ success: true, data: snap });
  }

  const all = await listDeviceSnapshots();
  return jsonResponse({ success: true, data: all });
}
