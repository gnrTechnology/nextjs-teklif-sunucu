import { listModuleOutputs, listModuleOutputsSummary, listHeartbeats } from "@/lib/db";
import ModulCiktilariClient from "@/app/components/ModulCiktilariClient";

export const dynamic = "force-dynamic";

export default async function ModulCiktilariPage() {
  const [outputs, summary, heartbeats] = await Promise.all([
    listModuleOutputs({ limit: 200 }),
    listModuleOutputsSummary(),
    listHeartbeats(),
  ]);

  const allMacs         = [...new Set(outputs.map((o) => o.mac))];
  const allModuleNames  = [...new Set(outputs.map((o) => o.moduleName))].sort();
  const allHostnames    = heartbeats.map((h) => ({ mac: h.mac, hostname: h.hostname ?? h.mac }));

  return (
    <ModulCiktilariClient
      initialOutputs={outputs}
      summary={summary}
      allMacs={allMacs}
      allModuleNames={allModuleNames}
      allHostnames={allHostnames}
    />
  );
}
