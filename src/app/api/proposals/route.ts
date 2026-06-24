import { jsonResponse } from "@/lib/api-response";
import { loadProposalsSummary } from "@/lib/proposals";

/** GET /api/proposals/ — öneri listesi + Neon modules tablosu karşılaştırması */
export async function GET() {
  const data = await loadProposalsSummary();
  return jsonResponse({ success: true, data });
}
