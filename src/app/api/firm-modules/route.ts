import fs from "fs";
import path from "path";
import { NextRequest } from "next/server";
import { jsonResponse, errorResponse } from "@/lib/api-response";
import type { FirmAutoModuleRecord } from "@/lib/types";

const FILE = path.join(process.cwd(), "data", "firm-auto-modules.json");

function readAll(): FirmAutoModuleRecord[] {
  if (!fs.existsSync(FILE)) return [];
  try {
    const raw = fs.readFileSync(FILE, "utf-8");
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : parsed.firms ?? [];
  } catch {
    return [];
  }
}

function writeAll(data: FirmAutoModuleRecord[]): void {
  fs.writeFileSync(FILE, JSON.stringify(data, null, 2), "utf-8");
}

/** GET /api/firm-modules */
export async function GET() {
  return jsonResponse({ success: true, data: readAll() });
}

/** POST /api/firm-modules — yeni firma ekle */
export async function POST(request: NextRequest) {
  let body: { firmaAdi: string; description?: string; enabled?: boolean };
  try {
    body = await request.json();
  } catch {
    return errorResponse("Geçersiz JSON.", 400);
  }

  if (!body.firmaAdi?.trim()) return errorResponse("firmaAdi zorunludur.", 400);

  const all = readAll();
  const existing = all.find(
    (f) => f.firmaAdi.toLowerCase() === body.firmaAdi.trim().toLowerCase(),
  );
  if (existing) return errorResponse("Bu firma zaten mevcut.", 409);

  const newItem: FirmAutoModuleRecord = {
    firmaAdi: body.firmaAdi.trim(),
    description: body.description ?? "",
    enabled: body.enabled ?? true,
    onExcelOpen: { enabled: true, modules: [] },
  };

  all.push(newItem);
  writeAll(all);
  return jsonResponse({ success: true, data: newItem }, 201);
}
