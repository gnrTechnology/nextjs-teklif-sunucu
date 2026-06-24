export const dynamic = "force-dynamic";

import { listFirmAutoModules } from "@/lib/firm-auto-modules";
import { listDbModules } from "@/lib/db";
import FirmaModulleriClient from "@/app/components/FirmaModulleriClient";

export default async function FirmaModulleriPage() {
  const [items, allModules] = await Promise.all([
    listFirmAutoModules(),
    listDbModules(),
  ]);

  const moduleNames = allModules.filter((m) => m.active !== false).map((m) => m.methodName);
  const moduleCategories: Record<string, string> = {};
  allModules.forEach((m) => { moduleCategories[m.methodName] = m.category ?? "genel"; });

  return (
    <FirmaModulleriClient
      initial={items}
      allModuleNames={moduleNames}
      moduleCategories={moduleCategories}
    />
  );
}
