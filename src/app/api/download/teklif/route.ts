import fs from "fs";
import path from "path";
import { NextRequest, NextResponse } from "next/server";
import { errorResponse } from "@/lib/api-response";
import { isLicensed } from "@/lib/db";
import type { DownloadPostBody } from "@/lib/types";

const ADDIN_FILE = path.join(process.cwd(), "data", "files", "teklif.xlam");

/**
 * VBA: LisansKontrolVeGuncelleme()
 * POST http://host:3000/api/download/teklif
 * Body: { "macAdresi": "AA:BB:CC:DD:EE:FF" }
 * Başarı: binary .xlam dosyası (200)
 */
export async function POST(request: NextRequest) {
  let body: DownloadPostBody;

  try {
    body = (await request.json()) as DownloadPostBody;
  } catch {
    return errorResponse("Geçersiz JSON gövdesi.", 400);
  }

  if (!body.macAdresi || body.macAdresi.trim() === "") {
    return errorResponse("macAdresi alanı zorunludur.", 400);
  }

  if (!isLicensed(body.macAdresi)) {
    return errorResponse(
      "Bu MAC adresi için aktif lisans bulunamadı. Önce lisans kaydı yapılmalıdır.",
      403,
    );
  }

  if (!fs.existsSync(ADDIN_FILE)) {
    return errorResponse(
      "Sunucuda eklenti dosyası (data/files/teklif.xlam) bulunamadı.",
      404,
    );
  }

  const fileBuffer = fs.readFileSync(ADDIN_FILE);

  return new NextResponse(fileBuffer, {
    status: 200,
    headers: {
      "Content-Type": "application/vnd.ms-excel.addin.macroEnabled",
      "Content-Disposition": 'attachment; filename="teklif.xlam"',
      "Content-Length": String(fileBuffer.length),
    },
  });
}
