export const dynamic = "force-dynamic";

import { listUnifiedActivity } from "@/lib/activity";
import LoglarClient from "@/app/components/LoglarClient";

export default async function LoglarPage() {
  const items = await listUnifiedActivity({ limit: 400 });
  return <LoglarClient initial={items} />;
}
