/**
 * Bozuk Debug.Print / string birlestirme taramasi (MsgBox fix yan etkisi)
 */
import fs from "fs";
import { neon } from "@neondatabase/serverless";

const sql = neon(
  fs.readFileSync(".env.local", "utf8").match(/^DATABASE_URL=(.+)$/m)[1].trim(),
);
const rows = await sql`SELECT method_name, code FROM modules`;

const patterns = [
  { re: /Debug\.Print\s+"[^"]*\bhttp\.Status\b[^"]*"/i, label: "literal http.Status in string" },
  { re: /Debug\.Print\s+"[^"]*\bErr\.Description\b[^"]*"/i, label: "literal Err.Description" },
  { re: /Debug\.Print\s+"[^"]*\bcount\b[^"]*"/i, label: "literal count in Debug.Print" },
  { re: /Debug\.Print\s+"[^"]*\bmodName\b[^"]*"/i, label: "literal modName" },
  { re: /Debug\.Print\s+"[^"]*\bprocName\b[^"]*"/i, label: "literal procName" },
  { re: /\$\{WS\}/, label: "template ${WS}" },
  { re: /sheetNote|WinAPI Declare; bkz/, label: "stub module" },
];

for (const { re, label } of patterns) {
  const hits = rows.filter((r) => re.test(r.code));
  console.log(`${label}: ${hits.length}`);
  hits.slice(0, 8).forEach((h) => console.log("  -", h.method_name));
}

// Modules missing proper exit
const noExit = rows.filter(
  (r) =>
    /Public Function DynamicFunc/i.test(r.code) &&
    !/Set DynamicFunc\s*=\s*Nothing/i.test(r.code),
);
console.log("\nDynamicFunc without Set DynamicFunc=Nothing:", noExit.length);
noExit.forEach((h) => console.log("  -", h.method_name));
