import { NextRequest } from "next/server";
import { errorResponse, jsonResponse } from "@/lib/api-response";
import { insertLog, upsertLicense } from "@/lib/db";
import type { LicensePostBody } from "@/lib/types";

// TODO [GÜVENLİK - YÜKSEK ÖNCELİK]: dosyaAdi ihlal kontrolü
// Her lisans başvurusunda body.dosyaAdi === "teklif.xlam" olup olmadığı kontrol edilecek.
// Farklıysa (dosya kopyalanıp yeniden adlandırılmış demektir):
//   1. licenses tablosunda license = 'false' yap
//   2. insertLog ile event_type='violation' yaz
//   3. Yanıtta { action: 'delete', targets: ['copy','addin'] } döndür
//   4. VBA bu yanıtı alınca hem kopya dosyayı hem teklif.xlam'ı siler
// Bakınız: getLicense.bas → SaveLicenseFromResponse + RegisterOrUpdate

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
