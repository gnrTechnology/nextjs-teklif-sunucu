import { listDbModules, ensureModulesTable } from "@/lib/db";
import ModullerClient from "@/app/components/ModullerClient";
import fs from "fs";
import path from "path";
import { upsertDbModule } from "@/lib/db";
import type { ModuleRecord } from "@/lib/types";

export const dynamic = "force-dynamic";

export default async function ModullerPage() {
  // Tablo oluştur, yoksa JSON'dan seed et
  await ensureModulesTable();
  let modules = await listDbModules();

  if (modules.length === 0) {
    const filePath = path.join(process.cwd(), "data", "modules.json");
    if (fs.existsSync(filePath)) {
      const list = JSON.parse(fs.readFileSync(filePath, "utf-8")) as ModuleRecord[];
      for (const item of list) {
        if (!item.methodName || !item.code) continue;
        await upsertDbModule({
          methodName: item.methodName,
          description: item.description ?? "",
          category: item.category ?? "genel",
          code: item.code,
          active: item.active ?? true,
        });
      }
      modules = await listDbModules();
    }
  }

  return <ModullerClient initial={modules} />;
}
