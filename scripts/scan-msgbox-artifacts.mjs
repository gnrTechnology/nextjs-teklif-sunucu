import fs from "fs";
import { neon } from "@neondatabase/serverless";
const sql = neon(
  fs.readFileSync(".env.local", "utf8").match(/^DATABASE_URL=(.+)$/m)[1].trim(),
);
const rows = await sql`SELECT method_name, code FROM modules`;
const patterns = [
  [/vbInformation|vbExclamation|vbCritical|vbYesNo/i, "MsgBox kalintisi (vb*)"],
  [/Debug\.Print[^\\n]*\),\s*vb/i, "bozuk Debug.Print parantez"],
  [/IIf\([^)]+\),\s*vb/i, "IIf+vb MsgBox kalintisi"],
  [/Chr\(10\)/i, "literal Chr(10) string icinde"],
];
for (const [re, label] of patterns) {
  const hits = rows.filter((r) => re.test(r.code));
  console.log(`${label}: ${hits.length}`);
  hits.slice(0, 15).forEach((h) => console.log("  -", h.method_name));
}
