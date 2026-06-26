/**
 * Uzak modul tikanma / runtime risk taramasi
 * node scripts/debug-modules-neon.mjs
 */
import fs from "fs";
import path from "path";
import { neon } from "@neondatabase/serverless";

const env = fs.readFileSync(".env.local", "utf8");
const sql = neon(env.match(/^DATABASE_URL=(.+)$/m)[1].trim());
const rows = await sql`SELECT method_name, code FROM modules ORDER BY method_name`;

const ISSUES = [];

function check(name, code) {
  const problems = [];
  const lines = code.split(/\r?\n/);

  if (!/Public Function DynamicFunc/i.test(code)) {
    problems.push("DynamicFunc yok");
  }

  if (/MsgBox/i.test(code)) problems.push("MsgBox kullanimi (tikanma riski)");

  if (/\bInputBox\b/i.test(code)) problems.push("InputBox (etkilesimli tikanma)");

  if (/Application\.Wait\b/i.test(code)) {
    const waits = (code.match(/Application\.Wait/g) || []).length;
    problems.push(`Application.Wait x${waits}`);
  }

  if (/\.Run\s+[^,]+,\s*0,\s*True\b/i.test(code)) {
    problems.push("Shell.Run sync=True (uzun blok)");
  }

  if (/DoEvents/i.test(code) && /While\b/i.test(code)) {
    problems.push("While+DoEvents dongusu");
  }

  if (/Set DynamicFunc\s*=\s*Nothing/i.test(code) === false && /Public Function DynamicFunc/i.test(code)) {
    problems.push("Set DynamicFunc = Nothing eksik");
  }

  if (/End Function/i.test(code) === false && /Public Function DynamicFunc/i.test(code)) {
    problems.push("End Function eksik");
  }

  // sheetNote stub — gercek is yapmaz ama takilmaz; yine de isaretle
  if (/WinAPI Declare; bkz dll-module-proposals/i.test(code)) {
    problems.push("Stub (sheetNote) — islev yok");
  }
  if (/param ile calisir; gerekirse zInternet/i.test(code)) {
    problems.push("Stub (sheetNote)");
  }

  // targetWb.Sheets(1) without On Error — common crash
  if (/targetWb\.Sheets\(1\)/i.test(code) && !/On Error/i.test(code)) {
    problems.push("targetWb.Sheets(1) On Error yok");
  }

  // ws.Columns.AutoFit on empty — slow but not hang
  if (/ws\.Cells\.ClearContents/i.test(code) && /AutoFit/i.test(code)) {
    // ok
  }

  // Infinite loop patterns
  if (/\bDo\s*$/im.test(code) || /\bDo\b[\s\S]*\bLoop\b/i.test(code)) {
    const doLoops = (code.match(/\bLoop\b/gi) || []).length;
    if (doLoops > 2) problems.push(`Coklu Do/Loop (${doLoops})`);
  }

  // Application.Interactive = False without restore
  if (/Application\.Interactive\s*=\s*False/i.test(code) && !/Application\.Interactive\s*=\s*True/i.test(code)) {
    problems.push("Interactive=False kalici olabilir");
  }

  // ScreenUpdating false without true
  if (/ScreenUpdating\s*=\s*False/i.test(code) && !/ScreenUpdating\s*=\s*True/i.test(code)) {
    problems.push("ScreenUpdating=False restore yok");
  }

  // Broken JSON escape in VBA
  if (/Replace\(s,\s*""""\s*,\s*"\\"""\)/.test(code)) {
    problems.push("Gecersiz JsonEsc");
  }

  // Template leftover
  if (/\$\{/.test(code)) problems.push("Template kalintisi");

  // Declare inside function
  let inFn = false;
  for (let i = 0; i < lines.length; i++) {
    const t = lines[i].trim();
    if (/^(Public |Private )?(Function|Sub)\s/i.test(t)) inFn = true;
    if (inFn && /^\s*Private\s+Declare\b/i.test(lines[i])) {
      problems.push(`Declare fonksiyon icinde satir ${i + 1}`);
    }
    if (/^End Function/i.test(t) || /^End Sub/i.test(t)) inFn = false;
  }

  // EnumWindows callback without End Function on DynamicFunc path
  if (/AddressOf\b/i.test(code) && !/End Function/i.test(code)) {
    problems.push("AddressOf callback — derleme riski");
  }

  // PowerShell -Command with broken quotes
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    if (/powershell/i.test(line)) {
      const dq = (line.match(/"/g) || []).length;
      if (dq % 2 !== 0) problems.push(`Satir ${i + 1}: PS tek tirnak`);
    }
  }

  // Missing Exit Function after error paths in long modules
  if (/GoTo Fail/i.test(code) && !/Fail:/i.test(code)) {
    problems.push("GoTo Fail etiketi yok");
  }

  if (/aktar/i.test(code)) {
    problems.push("kodda 'aktar' geciyor — kontrol");
  }

  if (problems.length) ISSUES.push({ name, problems });
}

for (const r of rows) check(r.method_name, r.code);

const outDir = path.join(process.cwd(), "data", "audit");
fs.mkdirSync(outDir, { recursive: true });
fs.writeFileSync(path.join(outDir, "debug-report.json"), JSON.stringify(ISSUES, null, 2));

const byType = {};
for (const item of ISSUES) {
  for (const p of item.problems) {
    byType[p] = (byType[p] || 0) + 1;
  }
}

console.log("Modul:", rows.length, "Sorunlu:", ISSUES.length);
console.log("\nSorun turleri:");
Object.entries(byType)
  .sort((a, b) => b[1] - a[1])
  .slice(0, 25)
  .forEach(([k, v]) => console.log(`  ${v}\t${k}`));

const stubs = ISSUES.filter((i) => i.problems.some((p) => p.includes("Stub")));
console.log("\nStub moduller:", stubs.length);

const hangRisk = ISSUES.filter((i) =>
  i.problems.some((p) =>
    /MsgBox|InputBox|Application\.Wait|sync=True|Interactive=False|ScreenUpdating/.test(p),
  ),
);
console.log("Tikanma riski:", hangRisk.length);
