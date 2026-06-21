import fs from "fs";
import path from "path";
import { NextResponse } from "next/server";
import { errorResponse } from "@/lib/api-response";

const REGISTRY_FILE = path.join(
  process.cwd(),
  "data",
  "registry-settings-full.json",
);

/**
 * VBA: ImportRegistrySettings modülü
 * GET http://host:3000/api/registry-settings
 */
export async function GET() {
  if (!fs.existsSync(REGISTRY_FILE)) {
    return errorResponse(
      "Sunucuda registry-settings-full.json bulunamadı.",
      404,
    );
  }

  const content = fs.readFileSync(REGISTRY_FILE, "utf-8");

  return new NextResponse(content, {
    status: 200,
    headers: {
      "Content-Type": "application/json; charset=utf-8",
    },
  });
}
