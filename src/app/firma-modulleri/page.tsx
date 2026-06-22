import { listFirmAutoModules } from "@/lib/firm-auto-modules";
import { listRemoteModuleNames } from "@/lib/modules";
import FirmaModulleriClient from "@/app/components/FirmaModulleriClient";

export default async function FirmaModulleriPage() {
  const [items, moduleNames] = await Promise.all([
    Promise.resolve(listFirmAutoModules()),
    listRemoteModuleNames(),
  ]);

  return <FirmaModulleriClient initial={items} allModuleNames={moduleNames} />;
}
