export const dynamic = "force-dynamic";

import { listDeviceSnapshots } from "@/lib/db";
import CihazlarClient from "@/app/components/CihazlarClient";

export default async function CihazlarPage() {
  const snapshots = await listDeviceSnapshots();
  return <CihazlarClient initial={snapshots} />;
}
