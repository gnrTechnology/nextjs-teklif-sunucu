import { NextRequest } from "next/server";
import { jsonResponse, errorResponse } from "@/lib/api-response";
import { updateClientCommand } from "@/lib/db";

/**
 * PATCH /api/commands/[id]
 * VBA modül tamamlandığında sonucu bildirir.
 * Body: { status: "done"|"error", result?: string, errorMsg?: string }
 */
export async function PATCH(
  request: NextRequest,
  { params }: { params: Promise<{ id: string }> },
) {
  const { id } = await params;
  const cmdId = Number(id);
  if (isNaN(cmdId)) return errorResponse("Geçersiz id.", 400);

  let body: { status: "done" | "error"; result?: string; errorMsg?: string };
  try { body = await request.json(); }
  catch { return errorResponse("Geçersiz JSON.", 400); }

  if (!["done", "error"].includes(body.status)) {
    return errorResponse("status: 'done' veya 'error' olmalı.", 400);
  }

  await updateClientCommand(cmdId, {
    status: body.status,
    result: body.result ?? null,
    errorMsg: body.errorMsg ?? null,
  });

  return jsonResponse({ success: true });
}
