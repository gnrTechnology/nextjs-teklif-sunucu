import { NextRequest } from "next/server";
import { jsonResponse, errorResponse } from "@/lib/api-response";
import { listDeviceSnapshotHistory } from "@/lib/db";

/**
 * GET /api/device-info/history?mac=...&limit=
 */
export async function GET(request: NextRequest) {
  const sp = new URL(request.url).searchParams;
  const mac = sp.get("mac");
  if (!mac?.trim()) return errorResponse("mac zorunludur.", 400);
  const limit = sp.get("limit") ? Number(sp.get("limit")) : 15;

  const data = await listDeviceSnapshotHistory(mac.trim(), limit);
  return jsonResponse({ success: true, data });
}
