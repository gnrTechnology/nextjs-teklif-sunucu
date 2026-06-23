import { NextRequest } from "next/server";
import { jsonResponse } from "@/lib/api-response";
import { listUnifiedActivity } from "@/lib/activity";
import type { ActivityCategory } from "@/lib/types";

/** GET /api/activity/?category=&limit=&mac= */
export async function GET(request: NextRequest) {
  const sp = new URL(request.url).searchParams;
  const category = (sp.get("category") ?? "all") as ActivityCategory;
  const limit = sp.get("limit") ? Number(sp.get("limit")) : 400;
  const mac = sp.get("mac") ?? undefined;

  const data = await listUnifiedActivity({ category, limit, mac });
  return jsonResponse({ success: true, data, count: data.length });
}
