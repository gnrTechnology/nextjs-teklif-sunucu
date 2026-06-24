export const dynamic = "force-dynamic";

import { listHeartbeats } from "@/lib/db";
import HeartbeatsClient from "@/app/components/HeartbeatsClient";
import { Suspense } from "react";

export default async function HeartbeatsPage() {
  const rows = await listHeartbeats();
  return (
    <Suspense>
      <HeartbeatsClient initial={rows} />
    </Suspense>
  );
}
