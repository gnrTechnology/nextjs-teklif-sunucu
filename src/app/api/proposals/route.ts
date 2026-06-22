import { jsonResponse } from "@/lib/api-response";
import { loadProposalsSummary } from "@/lib/proposals";

/** GET /api/proposals/ — öneri listesi + modules.json karşılaştırması */
export async function GET() {
  const data = loadProposalsSummary();
  return jsonResponse({ success: true, data });
}
