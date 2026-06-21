import { NextRequest } from "next/server";
import { errorResponse, jsonResponse } from "@/lib/api-response";
import { upsertLicense } from "@/lib/db";
import type { LicensePostBody } from "@/lib/types";

/**
 * VBA: RegisterLicense() ve PostDataToServer()
 * POST http://host:3000/api/license/
 */
export async function POST(request: NextRequest) {
  let body: LicensePostBody;

  try {
    const text = await request.text();
    if (!text || text.trim() === "") {
      return errorResponse("Boş istek gövdesi.", 400);
    }
    body = JSON.parse(text) as LicensePostBody;
  } catch (err) {
    console.error("[POST /api/license] JSON parse hatası:", err);
    return errorResponse("Geçersiz JSON gövdesi.", 400);
  }

  if (!body.macAdresi || body.macAdresi.trim() === "") {
    return errorResponse("macAdresi alanı zorunludur.", 400);
  }

  try {
    const { record, existed } = await upsertLicense(body);

    // VBA RegisterLicense 201 bekliyor; PostDataToServer 200 veya 201 kabul ediyor
    const status = existed ? 200 : 201;

    return jsonResponse(
      {
        success: true,
        message: existed ? "Lisans kaydı güncellendi." : "Yeni lisans kaydı oluşturuldu.",
        data: {
          macAdresi: record.macAdresi,
          license: record.license,
          firmaAdi: record.firmaAdi ?? null,
          userAdi: record.userAdi ?? null,
          dosyaAdi: record.dosyaAdi ?? null,
          ipAdresi: record.ipAdresi ?? null,
        },
      },
      status,
    );
  } catch (err) {
    console.error("[POST /api/license] DB hatası:", err);
    return errorResponse(
      "Sunucu hatası. Lütfen tekrar deneyin.",
      500,
      { detail: err instanceof Error ? err.message : String(err) },
    );
  }
}
