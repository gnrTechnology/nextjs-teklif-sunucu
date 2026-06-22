import { listHeartbeats } from "@/lib/db";
import HeartbeatsClient from "@/app/components/HeartbeatsClient";

export default async function HeartbeatsPage() {
  const rows = await listHeartbeats();
  return <HeartbeatsClient initial={rows} />;
}
