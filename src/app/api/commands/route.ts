import { NextRequest } from "next/server";
import { jsonResponse, errorResponse } from "@/lib/api-response";
import { createClientCommand, listClientCommands, deleteClientCommand } from "@/lib/db";

/** GET /api/commands?mac=&status=&limit= */
export async function GET(request: NextRequest) {
  const sp = new URL(request.url).searchParams;
  const mac    = sp.get("mac")    ?? undefined;
  const status = sp.get("status") ?? undefined;
  const limit  = sp.get("limit")  ? Number(sp.get("limit")) : 200;
  const data = await listClientCommands({ mac, status, limit });
  return jsonResponse({ success: true, data });
}

/** POST /api/commands — dashboard'dan istemciye komut gönder */
export async function POST(request: NextRequest) {
  let body: { mac: string; moduleName: string; param?: string; createdBy?: string };
  try { body = await request.json(); }
  catch { return errorResponse("Geçersiz JSON.", 400); }
  if (!body.mac?.trim())        return errorResponse("mac zorunludur.", 400);
  if (!body.moduleName?.trim()) return errorResponse("moduleName zorunludur.", 400);
  const cmd = await createClientCommand({
    mac: body.mac.trim(),
    moduleName: body.moduleName.trim(),
    param: body.param ?? null,
    createdBy: body.createdBy ?? "dashboard",
  });
  return jsonResponse({ success: true, data: cmd }, 201);
}

/** DELETE /api/commands?id= */
export async function DELETE(request: NextRequest) {
  const id = new URL(request.url).searchParams.get("id");
  if (!id) return errorResponse("id zorunludur.", 400);
  await deleteClientCommand(Number(id));
  return jsonResponse({ success: true });
}
