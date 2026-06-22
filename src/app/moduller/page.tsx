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

  // JSON'daki eksik modülleri DB'ye ekle
  const filePath = path.join(process.cwd(), "data", "modules.json");
  if (fs.existsSync(filePath)) {
    const list = JSON.parse(fs.readFileSync(filePath, "utf-8")) as ModuleRecord[];
    const existing = new Set(modules.map((m) => m.methodName.toLowerCase()));
    const newItems = list.filter(
      (item) => item.methodName && item.code && !existing.has(item.methodName.toLowerCase()),
    );
    if (newItems.length > 0) {
      for (const item of newItems) {
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
