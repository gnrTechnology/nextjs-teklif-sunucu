import { jsonResponse, errorResponse } from "@/lib/api-response";
import { listModuleOutputsSummary } from "@/lib/db";

/**
 * GET /api/module-output/summary
 * Her modülün toplam çalıştırma sayısı ve son çalışma zamanını döndürür.
 */
export async function GET() {
  try {
    const data = await listModuleOutputsSummary();
    return jsonResponse({ success: true, data });
  } catch (err) {
    return errorResponse(`Özet alınamadı: ${String(err)}`, 500);
  }
}
