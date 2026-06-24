/**
 * Neon modullerinde bilinen VBA syntax hatalarini duzeltir.
 * node scripts/fix-modules-neon.mjs
 * node scripts/fix-modules-neon.mjs CollectDeviceInfoServer
 */
import fs from "fs";
import { neon } from "@neondatabase/serverless";

const env = fs.readFileSync(".env.local", "utf8");
const url = env.match(/^DATABASE_URL=(.+)$/m)[1].trim();
const sql = neon(url);

const filter = process.argv[2];

/** @param {string} code */
function applyFixes(code) {
  let next = code;
  const fixes = [];

  // VBA'da "\"" gecersiz — derleme hatasi verir. Chr(34) kullan.
  const badEsc = /s = Replace\(s, """", "\\"""\)/g;
  if (badEsc.test(next)) {
    next = next.replace(badEsc, 's = Replace(s, """", "\\" & Chr(34))');
    fixes.push("EscJson cift tirnak escape");
  }

  return { code: next, fixes };
}

const query = filter
  ? sql`SELECT method_name, code FROM modules WHERE method_name = ${filter}`
  : sql`SELECT method_name, code FROM modules ORDER BY method_name`;

const rows = await query;
let updated = 0;

for (const row of rows) {
  const { code, fixes } = applyFixes(row.code);
  if (!fixes.length) continue;

  await sql`
    UPDATE modules
    SET code = ${code}, updated_at = NOW()
    WHERE method_name = ${row.method_name}
  `;
  updated++;
  console.log(`${row.method_name}: ${fixes.join(", ")}`);
}

console.log(`\nTaranan: ${rows.length}, guncellenen: ${updated}`);
