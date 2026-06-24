import { listDbModules, ensureModulesTable } from "@/lib/db";
import ModullerClient from "@/app/components/ModullerClient";

export const dynamic = "force-dynamic";

export default async function ModullerPage() {
  await ensureModulesTable();
  const modules = await listDbModules();
  return <ModullerClient initial={modules} />;
}
