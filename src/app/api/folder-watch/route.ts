import { NextRequest } from "next/server";
import { jsonResponse, errorResponse } from "@/lib/api-response";
import {
  insertFolderWatchEvent,
  listFolderWatchEvents,
  deleteFolderWatchEvent,
} from "@/lib/db";

/** GET /api/folder-watch/?mac=&limit= */
export async function GET(request: NextRequest) {
  const sp = new URL(request.url).searchParams;
  const mac = sp.get("mac") ?? undefined;
  const limit = sp.get("limit") ? Number(sp.get("limit")) : 200;
  const data = await listFolderWatchEvents({ mac, limit });
  return jsonResponse({ success: true, data });
}

/** POST /api/folder-watch/ — VBA WatchFolderServer olay bildirimi */
export async function POST(request: NextRequest) {
  let body: {
    mac?: string;
    hostname?: string;
    folderPath?: string;
    eventType?: string;
    fileName?: string;
    filePath?: string;
    detail?: string;
  };
  try {
    body = await request.json();
  } catch {
    return errorResponse("Geçersiz JSON.", 400);
  }

  if (!body.mac?.trim()) return errorResponse("mac zorunludur.", 400);
  if (!body.folderPath?.trim()) return errorResponse("folderPath zorunludur.", 400);
  if (!body.eventType?.trim()) return errorResponse("eventType zorunludur.", 400);

  await insertFolderWatchEvent({
    mac: body.mac.trim(),
    hostname: body.hostname ?? null,
    folderPath: body.folderPath.trim(),
    eventType: body.eventType.trim(),
    fileName: body.fileName ?? null,
    filePath: body.filePath ?? null,
    detail: body.detail ?? null,
  });

  return jsonResponse({ success: true });
}

/** DELETE /api/folder-watch/?id= */
export async function DELETE(request: NextRequest) {
  const idStr = new URL(request.url).searchParams.get("id");
  if (!idStr) return errorResponse("id zorunludur.", 400);
  const id = Number(idStr);
  if (!Number.isFinite(id) || id <= 0) return errorResponse("Geçersiz id.", 400);
  const ok = await deleteFolderWatchEvent(id);
  if (!ok) return errorResponse("Kayıt bulunamadı.", 404);
  return jsonResponse({ success: true });
}
