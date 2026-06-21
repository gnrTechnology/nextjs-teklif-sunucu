import fs from "fs";
import path from "path";
import { NextResponse } from "next/server";
import { errorResponse } from "@/lib/api-response";

const IHLAL_FILE = path.join(process.cwd(), "data", "files", "ihlal.xlsm");

/**
 * GET /api/download/ihlal
 * İhlal dosyasını indirir — auth gerekmez.
 * getLicense.bas ihlal tespitinde bu endpoint'i çağırır.
 */
export async function GET() {
  if (!fs.existsSync(IHLAL_FILE)) {
    return errorResponse(
      "İhlal dosyası (data/files/ihlal.xlsm) sunucuda bulunamadı.",
      404,
    );
  }

  const fileBuffer = fs.readFileSync(IHLAL_FILE);

  return new NextResponse(fileBuffer, {
    status: 200,
    headers: {
      "Content-Type":
        "application/vnd.ms-excel.sheet.macroEnabled.12",
      "Content-Disposition": 'attachment; filename="ihlal.xlsm"',
      "Content-Length": String(fileBuffer.length),
    },
  });
}
