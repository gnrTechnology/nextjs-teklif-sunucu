import { NextRequest } from "next/server";
import { errorResponse, jsonResponse } from "@/lib/api-response";
import { getRemoteModuleCode } from "@/lib/modules";
import type { ModulePostBody } from "@/lib/types";

/**
 * VBA: RunRemoteCode(methodName)
 * POST http://host:3000/api/module
 * Body: { "methodName": "ShowMessage" }
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

  const code = getRemoteModuleCode(body.methodName.trim());

  if (!code) {
    return errorResponse(`Bilinmeyen methodName: ${body.methodName}`, 404);
  }

  return jsonResponse({
    success: true,
    code,
  });
}
