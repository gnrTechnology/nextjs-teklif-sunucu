import fs from "fs";
import { neon } from "@neondatabase/serverless";
const sql = neon(fs.readFileSync(".env.local","utf8").match(/^DATABASE_URL=(.+)$/m)[1].trim());
const rows = await sql`SELECT method_name, code FROM modules`;
const broken = [];
for (const r of rows) {
  const lines = r.code.split(/\r?\n/);
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    if (!/Debug\.Print/i.test(line)) continue;
    // Debug.Print with only string literal (no & concatenation) but contains variable-like tokens
    if (/Debug\.Print\s+"[^"]*"/i.test(line) && !/&/.test(line)) {
      if (/\b(http\.|Err\.|count|modName|procName|fp|rc|r-1|satır|satir)\b/i.test(line)) {
        broken.push({ name: r.method_name, line: i + 1, text: line.trim() });
      }
    }
  }
}
console.log("Supheli Debug.Print:", broken.length);
broken.forEach((b) => console.log(`${b.name}:${b.line} ${b.text}`));
