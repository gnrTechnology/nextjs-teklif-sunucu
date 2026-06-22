import { NextRequest } from "next/server";
import { errorResponse, jsonResponse } from "@/lib/api-response";
import { getRemoteModuleCode } from "@/lib/modules";
import { incrementModuleRunCount } from "@/lib/db";
import type { ModulePostBody } from "@/lib/types";

/**
 * VBA: zInternet.RunRemoteCode "MethodName"
 * POST /api/module
 * Body: { "methodName": "getLicense" }
 * Yanıt: { "code": "Public Function DynamicFunc(...)" }
 */
export async function POST(request: NextRequest) {
  let body: ModulePostBody;

  try {
    body = (await request.json()) as ModulePostBody;
  } catch {
    return errorResponse("Geçersiz JSON gövdesi.", 400);
  }

  if (!body.methodName || body.methodName.trim() === "") {
    return errorResponse("methodName alanı zorunludur.", 400);
  }

  const name = body.methodName.trim();
  const code = await getRemoteModuleCode(name);

  if (!code) {
    return errorResponse(`Bilinmeyen metodAdı: ${body.methodName}`, 404);
  }

  /* Async counter — yanıtı geciktirmez */
  incrementModuleRunCount(name).catch(() => {});

  return jsonResponse({ success: true, code });
}
