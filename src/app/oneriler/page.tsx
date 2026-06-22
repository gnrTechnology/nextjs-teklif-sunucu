import { loadProposalsSummary } from "@/lib/proposals";
import OnerilerClient from "@/app/components/OnerilerClient";

export const dynamic = "force-dynamic";

export default function OnerilerPage() {
  const summary = loadProposalsSummary();
  return <OnerilerClient summary={summary} />;
}
