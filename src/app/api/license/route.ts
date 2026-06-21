import { NextRequest } from "next/server";
import { errorResponse, jsonResponse } from "@/lib/api-response";
import { getLicenseByMac, upsertLicense } from "@/lib/db";
import type { LicensePostBody } from "@/lib/types";

/**
 * VBA: RegisterLicense() ve PostDataToServer()
 * POST http://host:3000/api/license/
 */
export async function POST(request: NextRequest) {
  let body: LicensePostBody;

  try {
    body = (await request.json()) as LicensePostBody;
  } catch {
    return errorResponse("Geçersiz JSON gövdesi.", 400);
  }

  if (!body.macAdresi || body.macAdresi.trim() === "") {
    return errorResponse("macAdresi alanı zorunludur.", 400);
  }

  const existed = Boolean(getLicenseByMac(body.macAdresi));
  const record = upsertLicense(body);

  // VBA RegisterLicense 201 bekliyor; PostDataToServer 200 veya 201 kabul ediyor
  const status = existed ? 200 : 201;

  return jsonResponse(
    {
      success: true,
      message: existed
        ? "Lisans kaydı güncellendi."
        : "Yeni lisans kaydı oluşturuldu.",
      data: {
        macAdresi: record.macAdresi,
        license: record.license,
        firmaAdi: record.firmaAdi ?? null,
        dosyaAdi: record.dosyaAdi ?? null,
        ipAdresi: record.ipAdresi ?? null,
      },
    },
    status,
  );
}
