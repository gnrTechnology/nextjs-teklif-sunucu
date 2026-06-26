/**
 * Neon modullerinde Declare fonksiyon icindeyse dll-templates ile degistirir.
 * node scripts/fix-declare-neon.mjs
 */
import fs from "fs";
import path from "path";
import { neon } from "@neondatabase/serverless";

const ROOT = process.cwd();
const templatesPath = path.join(ROOT, "data", "dll-templates.json");
if (!fs.existsSync(templatesPath)) {
  console.error("Once: node scripts/generate-planned-remaining.mjs --export-dll-json");
  process.exit(1);
}
const templates = JSON.parse(fs.readFileSync(templatesPath, "utf8"));

const env = fs.readFileSync(".env.local", "utf8");
const url = env.match(/^DATABASE_URL=(.+)$/m)[1].trim();
const sql = neon(url);

function hasDeclareInFunction(code) {
  const lines = code.split(/\r?\n/);
  let inFunction = false;
  for (const line of lines) {
    const trimmed = line.trim();
    if (/^(Public |Private )?(Function|Sub)\s/i.test(trimmed)) inFunction = true;
    if (inFunction && /^\s*Private\s+Declare\b/i.test(line)) return true;
    if (/^End Function/i.test(trimmed) || /^End Sub/i.test(trimmed)) inFunction = false;
  }
  return false;
}

const rows = await sql`SELECT method_name, code FROM modules`;
let fixed = 0;

for (const row of rows) {
  if (!hasDeclareInFunction(row.code)) continue;
  const tpl = templates[row.method_name];
  if (!tpl || !/Private Declare/i.test(tpl)) continue;
  await sql`
    UPDATE modules SET code = ${tpl}, updated_at = NOW()
    WHERE method_name = ${row.method_name}
  `;
  console.log("Declare duzeltildi:", row.method_name);
  fixed++;
}

console.log("Toplam Declare duzeltme:", fixed);
