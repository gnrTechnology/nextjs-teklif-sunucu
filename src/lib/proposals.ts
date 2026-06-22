import fs from "fs";
import path from "path";

export type ProposalStatus = "done" | "planned" | "blocked";

export type ProposalItem = {
  methodName: string;
  description: string;
  status: ProposalStatus;
  section: string;
  source: "vba" | "dll";
  number: string;
};

export type ProposalsSummary = {
  items: ProposalItem[];
  implementedInDb: string[];
  stats: {
    total: number;
    done: number;
    planned: number;
    blocked: number;
    inDb: number;
    missingFromDb: number;
    extraInDb: number;
  };
  bySection: Record<string, { done: number; planned: number; inDb: number }>;
};

function parseProposalTable(
  content: string,
  source: "vba" | "dll",
): ProposalItem[] {
  const items: ProposalItem[] = [];
  let section = "Genel";

  for (const line of content.split(/\r?\n/)) {
    if (line.startsWith("## ")) {
      section = line.replace(/^##\s+(\d+\.\s*)?/, "").trim();
      continue;
    }
    if (!line.startsWith("|")) continue;
    if (line.includes("---") || line.includes("MethodName")) continue;

    const cols = line.split("|").map((c) => c.trim()).filter(Boolean);
    if (cols.length < 4) continue;

    const number = cols[0];
    const methodName = cols[1];
    const description = cols[2];
    const statusCol = cols[3];

    if (!/^[A-Za-z][A-Za-z0-9_]*$/.test(methodName)) continue;

    let status: ProposalStatus = "planned";
    if (statusCol.includes("✅")) status = "done";
    else if (statusCol.includes("🔒") || statusCol.includes("⚠️")) status = "blocked";

    items.push({ methodName, description, status, section, source, number });
  }

  return items;
}

function loadImplementedModuleNames(): Set<string> {
  const filePath = path.join(process.cwd(), "data", "modules.json");
  if (!fs.existsSync(filePath)) return new Set();
  try {
    const list = JSON.parse(fs.readFileSync(filePath, "utf-8")) as { methodName?: string }[];
    return new Set(
      list.filter((m) => m.methodName).map((m) => m.methodName!.toLowerCase()),
    );
  } catch {
    return new Set();
  }
}

export function loadProposalsSummary(): ProposalsSummary {
  const root = process.cwd();
  const vbaPath = path.join(root, "data", "module-proposals.md");
  const dllPath = path.join(root, "data", "dll-module-proposals.md");

  const items: ProposalItem[] = [];
  if (fs.existsSync(vbaPath)) {
    items.push(...parseProposalTable(fs.readFileSync(vbaPath, "utf-8"), "vba"));
  }
  if (fs.existsSync(dllPath)) {
    items.push(...parseProposalTable(fs.readFileSync(dllPath, "utf-8"), "dll"));
  }

  const dbNames = loadImplementedModuleNames();
  const implementedInDb = Array.from(dbNames).sort();

  const proposalDoneNames = new Set(
    items.filter((i) => i.status === "done").map((i) => i.methodName.toLowerCase()),
  );

  let missingFromDb = 0;
  for (const name of proposalDoneNames) {
    if (!dbNames.has(name)) missingFromDb++;
  }

  const proposalAllNames = new Set(items.map((i) => i.methodName.toLowerCase()));
  let extraInDb = 0;
  for (const name of dbNames) {
    if (!proposalAllNames.has(name)) extraInDb++;
  }

  const bySection: ProposalsSummary["bySection"] = {};
  for (const item of items) {
    if (!bySection[item.section]) {
      bySection[item.section] = { done: 0, planned: 0, inDb: 0 };
    }
    const s = bySection[item.section];
    if (item.status === "done") s.done++;
    else s.planned++;
    if (dbNames.has(item.methodName.toLowerCase())) s.inDb++;
  }

  return {
    items,
    implementedInDb,
    stats: {
      total: items.length,
      done: items.filter((i) => i.status === "done").length,
      planned: items.filter((i) => i.status === "planned").length,
      blocked: items.filter((i) => i.status === "blocked").length,
      inDb: dbNames.size,
      missingFromDb,
      extraInDb,
    },
    bySection,
  };
}
