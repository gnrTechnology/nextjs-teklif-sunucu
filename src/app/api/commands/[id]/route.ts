import { NextRequest } from "next/server";
import { jsonResponse, errorResponse } from "@/lib/api-response";
import { updateClientCommand, updateCommandProgress } from "@/lib/db";

/**
 * PATCH /api/commands/[id]
 * VBA modül ilerlemesi veya tamamlanma bildirimi.
 * Body: { progressPct?, progressLabel? } veya { status: "done"|"error", result?, errorMsg? }
 */
export async function PATCH(
  request: NextRequest,
  { params }: { params: Promise<{ id: string }> },
) {
  const { id } = await params;
  const cmdId = Number(id);
  if (isNaN(cmdId)) return errorResponse("Geçersiz id.", 400);

  let body: {
    status?: "done" | "error";
    result?: string;
    errorMsg?: string;
    progressPct?: number;
    progressLabel?: string;
  };
  try { body = await request.json(); }
  catch { return errorResponse("Geçersiz JSON.", 400); }

  if (body.status) {
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

  if (body.progressPct == null || !body.progressLabel?.trim()) {
    return errorResponse("progressPct ve progressLabel gerekli (veya status).", 400);
  }

  await updateCommandProgress(cmdId, body.progressPct, body.progressLabel.trim());
  return jsonResponse({ success: true });
}
