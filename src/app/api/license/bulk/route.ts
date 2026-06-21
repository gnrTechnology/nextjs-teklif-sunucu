import { NextRequest } from "next/server";
import { errorResponse, jsonResponse } from "@/lib/api-response";
import { bulkToggleLicenses, insertLog } from "@/lib/db";

/**
 * PATCH /api/license/bulk
 * Body: { macs: string[], action: "activate" | "deactivate" }
 */
export async function PATCH(request: NextRequest) {
  let body: { macs: string[]; action: "activate" | "deactivate" };

  try {
    body = await request.json();
  } catch {
    return errorResponse("Geçersiz JSON gövdesi.", 400);
  }

  if (!Array.isArray(body.macs) || body.macs.length === 0) {
    return errorResponse("macs alanı zorunludur ve boş olamaz.", 400);
  }
  if (body.action !== "activate" && body.action !== "deactivate") {
    return errorResponse("action alanı 'activate' veya 'deactivate' olmalıdır.", 400);
  }

  const newValue = body.action === "activate" ? "true" : "false";
  const count = await bulkToggleLicenses(body.macs, newValue);

  await insertLog({
    eventType: body.action === "activate" ? "activate" : "deactivate",
    details: `Toplu işlem: ${body.macs.length} cihaz ${body.action === "activate" ? "aktifleştirildi" : "pasifleştirildi"}.`,
  });

  return jsonResponse({
    success: true,
    affected: count,
    action: body.action,
    message: `${body.macs.length} lisans ${body.action === "activate" ? "aktifleştirildi" : "pasifleştirildi"}.`,
  });
}
