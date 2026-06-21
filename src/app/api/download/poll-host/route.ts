import fs from "fs";
import path from "path";
import { NextResponse } from "next/server";

const POLL_HOST_FILE = path.join(process.cwd(), "data", "files", "TeklifPollHost.xlsm");

/**
 * GET /api/download/poll-host/
 * TeklifPollHost.xlsm dosyasini indirir (lisans polling icin).
 */
export async function GET() {
  if (!fs.existsSync(POLL_HOST_FILE)) {
    return NextResponse.json(
      {
        success: false,
        message:
          "Sunucuda TeklifPollHost.xlsm bulunamadi. scripts/create-poll-host.ps1 calistirin.",
      },
      { status: 404 },
    );
  }

  const fileBuffer = fs.readFileSync(POLL_HOST_FILE);

  return new NextResponse(fileBuffer, {
    status: 200,
    headers: {
      "Content-Type": "application/vnd.ms-excel.sheet.macroEnabled",
      "Content-Disposition": 'attachment; filename="TeklifPollHost.xlsm"',
      "Content-Length": String(fileBuffer.length),
    },
  });
}
