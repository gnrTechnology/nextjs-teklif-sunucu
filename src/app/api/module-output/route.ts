import { NextRequest } from "next/server";
import { jsonResponse, errorResponse } from "@/lib/api-response";
import { insertModuleOutput, listModuleOutputs } from "@/lib/db";

/**
 * POST /api/module-output
 * VBA herhangi bir modülün çıktısını sunucuya gönderir.
 * Body: { mac, moduleName, hostname?, firmaAdi?, output: { key: value } }
 */
export async function POST(request: NextRequest) {
  let body: {
    mac: string;
    moduleName: string;
    hostname?: string;
    firmaAdi?: string;
    output: Record<string, unknown>;
  };
  try { body = await request.json(); }
  catch { return errorResponse("Geçersiz JSON.", 400); }

  if (!body.mac?.trim())        return errorResponse("mac zorunludur.", 400);
  if (!body.moduleName?.trim()) return errorResponse("moduleName zorunludur.", 400);
  if (!body.output || typeof body.output !== "object") {
    return errorResponse("output objesi zorunludur.", 400);
  }

  await insertModuleOutput({
    mac: body.mac.trim(),
    moduleName: body.moduleName.trim(),
    hostname: body.hostname ?? null,
    firmaAdi: body.firmaAdi ?? null,
    output: body.output,
  });

  return jsonResponse({ success: true, message: "Modül çıktısı kaydedildi." });
}

/**
 * GET /api/module-output?mac=&moduleName=&limit=
 */
export async function GET(request: NextRequest) {
  const sp = new URL(request.url).searchParams;
  const mac        = sp.get("mac")        ?? undefined;
  const moduleName = sp.get("moduleName") ?? undefined;
  const limit      = sp.get("limit")      ? Number(sp.get("limit")) : 100;

  const data = await listModuleOutputs({ mac, moduleName, limit });
  return jsonResponse({ success: true, data });
}
