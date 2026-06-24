import { loadProposalsSummary } from "@/lib/proposals";
import OnerilerClient from "@/app/components/OnerilerClient";

export const dynamic = "force-dynamic";

export default async function OnerilerPage() {
  const summary = await loadProposalsSummary();
  return <OnerilerClient summary={summary} />;
}
