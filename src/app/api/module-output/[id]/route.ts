import { NextRequest } from "next/server";
import { jsonResponse, errorResponse } from "@/lib/api-response";
import { deleteModuleOutput } from "@/lib/db";

export async function DELETE(
  _req: NextRequest,
  { params }: { params: Promise<{ id: string }> }
) {
  const { id } = await params;
  const numId = Number(id);
  if (!numId || isNaN(numId)) return errorResponse("Geçersiz id.", 400);
  try {
    await deleteModuleOutput(numId);
    return jsonResponse({ success: true });
  } catch (err) {
    return errorResponse(`Silme hatası: ${String(err)}`, 500);
  }
}
