import { listClientCommands } from "@/lib/db";
import { listHeartbeats } from "@/lib/db";
import { listDbModules } from "@/lib/db";
import KomutlarClient from "@/app/components/KomutlarClient";

export default async function KomutlarPage() {
  const [commands, heartbeats, modules] = await Promise.all([
    listClientCommands({ limit: 200 }),
    listHeartbeats(),
    listDbModules(),
  ]);

  const allMacs = heartbeats.map((h) => h.mac);
  const allModuleNames = modules.filter((m) => m.active !== false).map((m) => m.methodName);

  return (
    <KomutlarClient
      initial={commands}
      allModuleNames={allModuleNames}
      allMacs={allMacs}
    />
  );
}
