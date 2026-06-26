/**
 * Neon modullerinde MsgBox -> Debug.Print
 * node scripts/fix-msgbox-neon.mjs
 * node scripts/fix-msgbox-neon.mjs --dry-run
 */
import fs from "fs";
import { neon } from "@neondatabase/serverless";

const dryRun = process.argv.includes("--dry-run");
const env = fs.readFileSync(".env.local", "utf8");
const url = env.match(/^DATABASE_URL=(.+)$/m)[1].trim();
const sql = neon(url);

/** @param {string} code @param {string} methodName */
function stripMsgBox(code, methodName) {
  if (!/MsgBox/i.test(code)) return { code, changed: false };

  let next = code;
  const tag = `[${methodName}]`;

  // MsgBox "text", vb..., "title"  (tek satir)
  next = next.replace(
    /MsgBox\s+((?:"[^"]*"|[^,\n]+)(?:\s*&\s*(?:"[^"]*"|[^,\n]+))*)\s*(?:,\s*[^,\n]+)?(?:\s*,\s*"[^"]*")?/gi,
    (match, textPart) => {
      const cleaned = textPart
        .replace(/\s*&\s*vbCrLf\s*&\s*/gi, " ")
        .replace(/\s*&\s*/g, " ")
        .replace(/"/g, "")
        .trim();
      return `Debug.Print "${tag} ${cleaned}"`;
    },
  );

  // If Not quiet Then MsgBox ...
  next = next.replace(
    /If\s+Not\s+quiet\s+Then\s+MsgBox\s+([^,\n]+(?:\s*&\s*[^,\n]+)*)/gi,
    (match, textPart) => {
      const cleaned = textPart.replace(/"/g, "").trim();
      return `If Not quiet Then Debug.Print "${tag} ${cleaned}"`;
    },
  );

  // If MsgBox(...) <> vbYes Then ...
  next = next.replace(
    /If\s+MsgBox\s*\(([^)]+(?:\([^)]*\)[^)]*)*)\)\s*<>\s*vbYes\s+Then\s+/gi,
    (match, inner) => {
      const cleaned = inner.replace(/"/g, "").replace(/\s+/g, " ").trim();
      return `Debug.Print "${tag} ${cleaned}"\r\n    If False Then `;
    },
  );

  // Dim x = MsgBox(...)
  next = next.replace(
    /Dim\s+(\w+)\s+As\s+Long\s*:\s*\1\s*=\s*MsgBox\s*\(([^)]+)\)/gi,
    (match, varName, inner) => {
      const cleaned = inner.replace(/"/g, "").trim();
      return `Dim ${varName} As Long : Debug.Print "${tag} ${cleaned}" : ${varName} = vbYes`;
    },
  );

  const changed = next !== code;
  return { code: next, changed };
}

const rows = await sql`SELECT method_name, code FROM modules ORDER BY method_name`;
let updated = 0;
const touched = [];

for (const row of rows) {
  const { code, changed } = stripMsgBox(row.code, row.method_name);
  if (!changed) continue;
  touched.push(row.method_name);
  if (!dryRun) {
    await sql`
      UPDATE modules SET code = ${code}, updated_at = NOW()
      WHERE method_name = ${row.method_name}
    `;
  }
  updated++;
}

console.log(dryRun ? "[dry-run] " : "", `Taranan: ${rows.length}, MsgBox giderilen: ${updated}`);
touched.forEach((n) => console.log(" -", n));
