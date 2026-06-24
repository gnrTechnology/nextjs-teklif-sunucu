import { NextRequest } from "next/server";
import { jsonResponse, errorResponse } from "@/lib/api-response";
import {
  deleteCommandTemplate,
  insertCommandTemplate,
  listCommandTemplates,
} from "@/lib/db";

export async function GET() {
  const data = await listCommandTemplates();
  return jsonResponse({ success: true, data });
}

export async function POST(request: NextRequest) {
  let body: { label?: string; moduleName?: string; param?: string | null };
  try {
    body = await request.json();
  } catch {
    return errorResponse("Geçersiz JSON.", 400);
  }
  if (!body.label?.trim()) return errorResponse("label zorunludur.", 400);
  if (!body.moduleName?.trim()) return errorResponse("moduleName zorunludur.", 400);

  const data = await insertCommandTemplate({
    label: body.label.trim(),
    moduleName: body.moduleName.trim(),
    param: body.param ?? null,
  });
  return jsonResponse({ success: true, data }, 201);
}

export async function DELETE(request: NextRequest) {
  const id = Number(new URL(request.url).searchParams.get("id"));
  if (!id) return errorResponse("id zorunludur.", 400);
  await deleteCommandTemplate(id);
  return jsonResponse({ success: true });
}
