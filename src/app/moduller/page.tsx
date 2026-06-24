import { Suspense } from "react";
import { listDbModules, ensureModulesTable } from "@/lib/db";
import ModullerClient from "@/app/components/ModullerClient";

export const dynamic = "force-dynamic";

export default async function ModullerPage() {
  await ensureModulesTable();
  const modules = await listDbModules();
  return (
    <Suspense>
      <ModullerClient initial={modules} />
    </Suspense>
  );
}
