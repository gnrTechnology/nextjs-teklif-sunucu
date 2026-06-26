/**
 * dll-module-proposals.md icindeki ✅ modulleri Declare sablonlariyla Neon'a yazar.
 */
import fs from "fs";
import path from "path";
import { execSync } from "child_process";
import { neon } from "@neondatabase/serverless";
import { CUSTOM } from "./generate-planned-remaining.mjs";

const dryRun = process.argv.includes("--dry-run");
const ROOT = process.cwd();

execSync("node scripts/generate-planned-remaining.mjs --export-dll-json", {
  cwd: ROOT,
  stdio: "inherit",
});

const templates = JSON.parse(
  fs.readFileSync(path.join(ROOT, "data", "dll-templates.json"), "utf8"),
);

const proposals = fs.readFileSync(
  path.join(ROOT, "data", "dll-module-proposals.md"),
  "utf8",
);
const approved = new Set();
for (const line of proposals.split(/\r?\n/)) {
  const m = line.match(/^\|\s*D\d+\s*\|\s*([A-Za-z][A-Za-z0-9_]*)\s*\|/);
  if (m && /✅/.test(line)) approved.add(m[1]);
}

const env = fs.readFileSync(".env.local", "utf8");
const url = env.match(/^DATABASE_URL=(.+)$/m)[1].trim();
const sql = neon(url);

const rows = await sql`SELECT method_name FROM modules`;
const inDb = new Set(rows.map((r) => r.method_name));

const targets = new Set([...approved, ...Object.keys(templates), ...inDb]);
let updated = 0;

for (const name of targets) {
  if (!approved.has(name)) continue;
  const code = templates[name] || CUSTOM[name];
  if (!code || !/Private Declare/i.test(code)) continue;
  if (!inDb.has(name) && dryRun) {
    console.log("[dry-run] eklenecek:", name);
    updated++;
    continue;
  }
  if (!inDb.has(name) && !dryRun) {
    await sql`
      INSERT INTO modules (method_name, description, category, active, code, created_at, updated_at)
      VALUES (${name}, ${name}, 'dll', true, ${code.trim() + "\n"}, NOW(), NOW())
    `;
    console.log("eklendi:", name);
    updated++;
    continue;
  }
  if (!inDb.has(name)) continue;
  if (dryRun) {
    console.log("[dry-run] guncellenecek:", name);
    updated++;
    continue;
  }
  await sql`
    UPDATE modules SET code = ${code.trim() + "\n"}, updated_at = NOW()
    WHERE method_name = ${name}
  `;
  console.log("guncellendi:", name);
  updated++;
}

console.log(
  dryRun ? "[dry-run] " : "",
  `DLL sync: ${updated} modul, sablon: ${Object.keys(templates).length}, onayli: ${approved.size}`,
);
