/**
 * MsgBox onay donusumunden kalan vb* ve If False Then kalintilarini temizler.
 */
import fs from "fs";
import { neon } from "@neondatabase/serverless";

const dryRun = process.argv.includes("--dry-run");
const sql = neon(
  fs.readFileSync(".env.local", "utf8").match(/^DATABASE_URL=(.+)$/m)[1].trim(),
);

function fixConfirmModules(code) {
  let c = code;
  c = c.replace(
    /Debug\.Print\s+"[^"]*vbYesNo[^"]*"\s*\r?\n\s*If False Then GoTo Done\s*\r?\n/gi,
    "",
  );
  c = c.replace(/Debug\.Print\s+"\[[^\]]+\][^"]*vbYesNo[^"]*"\s*\r?\n/gi, "");
  c = c.replace(/Debug\.Print\s+"\[[^\]]+\][^"]*vbExclamation[^"]*"\s*\r?\n/gi, "");
  c = c.replace(/\r?\n\s*If False Then GoTo Done\s*\r?\n/gi, "\n");
  return c;
}

function fixBrokenDebugPrint(code) {
  let c = code;
  const fixes = [
    [
      /Debug\.Print\s+"\[KillProcessByName\]\s*procName\s+sonlandırıldı\."/,
      'Debug.Print "[KillProcessByName] " & procName & " sonlandirildi."',
    ],
    [
      /Debug\.Print\s+"\[getLicense\]\s*Bağlantı hatası:\s*Err\.Description"/,
      'Debug.Print "[getLicense] Baglanti hatasi: " & Err.Description',
    ],
    [
      /Debug\.Print\s+"\[WriteTextFile\]\s*Dosya yazıldı:\s*fp"/,
      'Debug.Print "[WriteTextFile] Dosya yazildi: " & fp',
    ],
    [
      /Debug\.Print\s+"\[AppendToTextFile\]\s*Satır eklendi:\s*fp"/,
      'Debug.Print "[AppendToTextFile] Satir eklendi: " & fp',
    ],
    [
      /Debug\.Print\s+"\[WriteSheetToCsv\]\s*CSV kaydedildi:\s*fp"/,
      'Debug.Print "[WriteSheetToCsv] CSV kaydedildi: " & fp',
    ],
    [
      /Debug\.Print\s+"\[GetFolderSize\]\s*fp Chr\(10\) Toplam:\s*result"/,
      'Debug.Print "[GetFolderSize] " & fp & Chr(10) & "Toplam: " & result',
    ],
    [
      /Debug\.Print\s+"\[AutoArchiveOldRows\]\s*moved\s+satır arşivlendi \( archSheet \)\."/,
      'Debug.Print "[AutoArchiveOldRows] " & moved & " satir arsivlendi (" & archSheet & ")."',
    ],
  ];
  for (const [re, rep] of fixes) {
    c = c.replace(re, rep);
  }
  c = c.replace(/Debug\.Print([^,\n]+),vbExclamation:GoTo Done/gi, (m, inner) => {
    const tag = inner.match(/"(\[[^\]]+\])/)?.[1] || "[modul]";
    return `Debug.Print "${tag} Param eksik"\r\n        GoTo Done`;
  });
  return c;
}

const rows = await sql`SELECT method_name, code FROM modules`;
let n = 0;
for (const row of rows) {
  let fixed = fixConfirmModules(row.code);
  fixed = fixBrokenDebugPrint(fixed);
  if (fixed === row.code) continue;
  if (!dryRun) {
    await sql`UPDATE modules SET code = ${fixed}, updated_at = NOW() WHERE method_name = ${row.method_name}`;
  }
  console.log("fixed:", row.method_name);
  n++;
}
console.log(dryRun ? "[dry-run] " : "", "Toplam:", n);
