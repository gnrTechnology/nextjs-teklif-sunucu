import { NextRequest } from "next/server";
import { upsertHeartbeat } from "@/lib/db";
import { jsonResponse, errorResponse } from "@/lib/api-response";

/**
 * POST /api/heartbeat
 * Body: { mac, hostname, user, excelVersion, timestamp }
 */
export async function POST(request: NextRequest) {
  let body: {
    mac?: string;
    hostname?: string;
    user?: string;
    excelVersion?: string;
  };

  try {
    body = await request.json();
  } catch {
    return errorResponse("Geçersiz JSON.", 400);
  }

  if (!body.mac?.trim()) return errorResponse("mac alanı zorunludur.", 400);

  const ip =
    request.headers.get("x-forwarded-for")?.split(",")[0]?.trim() ??
    request.headers.get("x-real-ip") ??
    null;

  await upsertHeartbeat({
    mac:          body.mac.trim(),
    hostname:     body.hostname ?? null,
    userName:     body.user    ?? null,
    excelVersion: body.excelVersion ?? null,
    ipAddress:    ip,
  });

  return jsonResponse({ success: true });
}

/** GET /api/heartbeat — dashboard'un otomatik yenileme için kullandığı endpoint */
export async function GET() {
  const { listHeartbeats } = await import("@/lib/db");
  const rows = await listHeartbeats();
  return jsonResponse({ success: true, data: rows });
}
