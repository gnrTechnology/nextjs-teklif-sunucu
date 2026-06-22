import { NextRequest } from "next/server";
import { jsonResponse, errorResponse } from "@/lib/api-response";
import { claimPendingCommand, normalizeMac } from "@/lib/db";

/**
 * GET /api/commands/pending/[mac]
 * VBA HeartbeatPing her çalışmada bunu sorgular.
 * Bekleyen komut varsa "running" durumuna çeker ve döner.
 * Yoksa { success:true, data:null } döner.
 */
export async function GET(
  _request: NextRequest,
  { params }: { params: Promise<{ mac: string }> },
) {
  const { mac } = await params;
  const decoded = decodeURIComponent(mac);
  if (!decoded || decoded.length < 10) return errorResponse("Geçersiz MAC.", 400);
  const cmd = await claimPendingCommand(normalizeMac(decoded));
  return jsonResponse({ success: true, data: cmd });
}
