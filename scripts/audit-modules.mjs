/**
 * Neon modullerini VBA syntax acisindan tarar.
 * node scripts/audit-modules.mjs
 * node scripts/audit-modules.mjs CollectDeviceInfoServer
 */
import fs from "fs";
import path from "path";
import { neon } from "@neondatabase/serverless";

const env = fs.readFileSync(".env.local", "utf8");
const url = env.match(/^DATABASE_URL=(.+)$/m)[1].trim();
const sql = neon(url);

const filter = process.argv[2];

const ISSUES = [];

function checkVba(name, code) {
  const problems = [];

  if (!code.includes("Public Function DynamicFunc") && !code.includes("Public Sub ")) {
    problems.push("DynamicFunc veya Public Sub yok");
  }

  // Declare inside function body (invalid VBA)
  const lines = code.split(/\r?\n/);
  let inFunction = false;
  let depth = 0;
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const trimmed = line.trim();
    if (/^(Public |Private )?(Function|Sub)\s/i.test(trimmed)) {
      inFunction = true;
      depth = 0;
    }
    if (inFunction && /^\s*Private\s+Declare\b/i.test(line)) {
      problems.push(`Satir ${i + 1}: Declare fonksiyon icinde (modul seviyesine tasinmali)`);
    }
    if (inFunction && /^\s*#If\b/i.test(trimmed)) {
      problems.push(`Satir ${i + 1}: #If fonksiyon icinde olabilir — kontrol edin`);
    }
    if (/^End Function/i.test(trimmed) || /^End Sub/i.test(trimmed)) inFunction = false;

    // Gecersiz JSON escape: VBA'da "\"" derlenmez
    if (/s = Replace\(s, """", "\\"""\)/.test(line)) {
      problems.push(`Satir ${i + 1}: Gecersiz VBA escape (\\""") — Chr(34) kullanin`);
    }

    // Unclosed strings — powershell "" kalibi false positive uretir, atla
    const isPsRunLine =
      /powershell\s+-NoProfile/i.test(line) || /sh\.Run\(/i.test(line);
    const dq = (line.match(/"/g) || []).length;
    if (dq % 2 !== 0 && !trimmed.startsWith("'") && !isPsRunLine) {
      problems.push(`Satir ${i + 1}: Tek sayida cift tirnak`);
    }
  }

  // Private Declare after Public Function started without module-level block
  if (/Public Function[\s\S]*\n\s*Private Declare/m.test(code)) {
    problems.push("Declare Public Function sonrasinda — derlenmez");
  }

  // Invalid #If inside wrap-only stubs
  if (code.includes("Private Declare") && code.includes("Public Function")) {
    const declareIdx = code.indexOf("Private Declare");
    const funcIdx = code.indexOf("Public Function");
    if (declareIdx > funcIdx) {
      problems.push("Declare, Public Function'dan sonra geliyor");
    }
  }

  // Common generator bug: ${WS} literal left in code
  if (code.includes("${WS}") || code.includes("${")) {
    problems.push("Template kalintisi (${...})");
  }

  // End Function count
  const funcStarts = (code.match(/Public Function DynamicFunc/g) || []).length;
  const funcEnds = (code.match(/^End Function/im) || []).length;
  if (funcStarts > 0 && funcEnds !== funcStarts) {
    problems.push(`End Function sayisi uyumsuz (${funcEnds}/${funcStarts})`);
  }

  if (problems.length) ISSUES.push({ name, problems });
}

const query = filter
  ? sql`SELECT method_name, code FROM modules WHERE method_name = ${filter}`
  : sql`SELECT method_name, code FROM modules ORDER BY method_name`;

const rows = await query;
for (const r of rows) checkVba(r.method_name, r.code);

const outDir = path.join(process.cwd(), "data", "audit");
fs.mkdirSync(outDir, { recursive: true });

if (filter && rows[0]) {
  fs.writeFileSync(path.join(outDir, `${filter}.bas`), rows[0].code, "utf8");
  console.log("written:", path.join(outDir, `${filter}.bas`));
}

console.log(`Scanned: ${rows.length}, issues: ${ISSUES.length}`);
for (const item of ISSUES.slice(0, 50)) {
  console.log(`\n${item.name}:`);
  item.problems.forEach((p) => console.log(`  - ${p}`));
}
if (ISSUES.length > 50) console.log(`\n... ve ${ISSUES.length - 50} modul daha`);

fs.writeFileSync(
  path.join(outDir, "issues.json"),
  JSON.stringify(ISSUES, null, 2),
  "utf8",
);
