export const dynamic = "force-dynamic";

import { listFolderWatchEvents, listHeartbeats } from "@/lib/db";
import KlasorIzlemeClient from "@/app/components/KlasorIzlemeClient";

export default async function KlasorIzlemePage() {
  const [events, heartbeats] = await Promise.all([
    listFolderWatchEvents({ limit: 200 }),
    listHeartbeats(),
  ]);
  return (
    <KlasorIzlemeClient
      initial={events}
      heartbeats={heartbeats.map((h) => ({ mac: h.mac, hostname: h.hostname }))}
    />
  );
}
