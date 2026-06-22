import fs from "fs";
import path from "path";
import { NextRequest, NextResponse } from "next/server";
import { errorResponse } from "@/lib/api-response";

const AGENT_DIR = path.join(process.cwd(), "data", "agent");

/**
 * GET /api/agent/download?arch=x64|x86&file=dll|exe
 * TeklifAgent COM DLL veya exe indirir.
 */
export async function GET(request: NextRequest) {
  const arch = request.nextUrl.searchParams.get("arch") ?? "x64";
  const file = request.nextUrl.searchParams.get("file") ?? "dll";

  let fileName: string;
  if (file === "exe") {
    fileName = "TeklifAgent.exe";
  } else if (arch === "x86") {
    fileName = "TeklifAgent.Com.x86.dll";
  } else {
    fileName = "TeklifAgent.Com.dll";
  }

  const filePath = path.join(AGENT_DIR, fileName);
  if (!fs.existsSync(filePath)) {
    return errorResponse(
      `Agent dosyasi bulunamadi: ${fileName}. Sunucuda teklif-agent/build-agent.ps1 calistirin.`,
      404,
    );
  }

  const buf = fs.readFileSync(filePath);
  const contentType =
    file === "exe"
      ? "application/octet-stream"
      : "application/x-msdownload";

  return new NextResponse(buf, {
    status: 200,
    headers: {
      "Content-Type": contentType,
      "Content-Disposition": `attachment; filename="${fileName === "TeklifAgent.Com.x86.dll" ? "TeklifAgent.Com.dll" : fileName}"`,
      "Content-Length": String(buf.length),
      "Cache-Control": "no-cache",
    },
  });
}
