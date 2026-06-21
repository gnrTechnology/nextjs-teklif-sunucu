import { NextRequest } from "next/server";
import { errorResponse, jsonResponse } from "@/lib/api-response";
import { insertLog, toggleLicense, upsertLicense } from "@/lib/db";
import type { LicensePostBody } from "@/lib/types";

const ALLOWED_FILENAME = "teklif.xlam";

/**
 * VBA: RegisterLicense() ve PostDataToServer()
 * POST http://host:3000/api/license/
 *
 * Güvenlik: dosyaAdi "teklif.xlam" değilse → ihlal tespiti.
 *   - Lisans zorla false yapılır
 *   - violation logu yazılır
 *   - Yanıtta { ihlal: true } döner → VBA ihlal.xlsm indirir ve çalıştırır
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

  // İhlal tespiti: dosyaAdi gönderilmiş ama teklif.xlam değil
  const isViolation =
    typeof body.dosyaAdi === "string" &&
    body.dosyaAdi.trim() !== "" &&
    body.dosyaAdi.trim().toLowerCase() !== ALLOWED_FILENAME.toLowerCase();

  try {
    const { record, existed } = await upsertLicense(body);

    if (isViolation) {
      // Lisansı zorla kapat
      await toggleLicense(record.macAdresi, "false");

      await insertLog({
        macAdresi: record.macAdresi,
        firmaAdi: record.firmaAdi,
        userAdi: record.userAdi,
        dosyaAdi: body.dosyaAdi,
        ipAdresi: record.ipAdresi,
        eventType: "violation",
        details: `İzinsiz kopya tespit edildi: "${body.dosyaAdi}". Lisans devre dışı bırakıldı.`,
      });

      return jsonResponse(
        {
          success: true,
          message: "Lisans kaydı oluşturuldu ancak ihlal tespit edildi.",
          ihlal: true,
          kopyaDosyaAdi: body.dosyaAdi,
          data: {
            macAdresi: record.macAdresi,
            license: "false",
            firmaAdi: record.firmaAdi ?? null,
            userAdi: record.userAdi ?? null,
            dosyaAdi: body.dosyaAdi ?? null,
          },
        },
        200,
      );
    }

    await insertLog({
      macAdresi: record.macAdresi,
      firmaAdi: record.firmaAdi,
      userAdi: record.userAdi,
      dosyaAdi: record.dosyaAdi,
      ipAdresi: record.ipAdresi,
      eventType: existed ? "update" : "register",
      details: existed ? "Mevcut kayıt güncellendi." : "Yeni lisans kaydı oluşturuldu.",
    });

    const status = existed ? 200 : 201;

    return jsonResponse(
      {
        success: true,
        message: existed ? "Lisans kaydı güncellendi." : "Yeni lisans kaydı oluşturuldu.",
        ihlal: false,
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
