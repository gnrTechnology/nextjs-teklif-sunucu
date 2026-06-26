/**
 * Neon modullerinde runtime/derleme hatalarini duzeltir.
 * node scripts/fix-runtime-neon.mjs
 * node scripts/fix-runtime-neon.mjs --dry-run
 */
import fs from "fs";
import path from "path";
import { neon } from "@neondatabase/serverless";

const dryRun = process.argv.includes("--dry-run");
const env = fs.readFileSync(".env.local", "utf8");
const sql = neon(env.match(/^DATABASE_URL=(.+)$/m)[1].trim());

/** Bilinen satir bazli duzeltmeler (MsgBox->Debug.Print bozulmasi) */
const EXACT_LINE_FIXES = {
  FileExists:
    '    Debug.Print "[FileExists] " & fp & Chr(10) & IIf(exists, "Evet", "Hayir")',
  ReadCsvToSheet:
    '    Debug.Print "[ReadCsvToSheet] " & (r - 1) & " satir aktarildi."',
  WriteTextFile: '    Debug.Print "[WriteTextFile] Dosya yazildi: " & fp',
  AppendToTextFile: '    Debug.Print "[AppendToTextFile] Satir eklendi: " & fp',
  WriteSheetToCsv: '    Debug.Print "[WriteSheetToCsv] CSV kaydedildi: " & fp',
  GetFolderSize:
    '    Debug.Print "[GetFolderSize] " & fp & Chr(10) & "Toplam: " & result',
  GetFileSize: {
    12: '    Debug.Print "[GetFileSize] " & fso.GetFileName(fp) & ": " & szStr',
  },
  KillProcessByName:
    '    Debug.Print "[KillProcessByName] " & procName & " sonlandirildi."',
  InjectVbaModule:
    '    Debug.Print "[InjectVbaModule] Modul enjekte edildi: " & modName',
  RemoveVbaModule: [
    '    Debug.Print "[RemoveVbaModule] Modul silindi: " & modName',
    '    Debug.Print "[RemoveVbaModule] Modul bulunamadi: " & modName',
  ],
  AutoArchiveOldRows:
    '    Debug.Print "[AutoArchiveOldRows] " & moved & " satir arsivlendi (" & archSheet & ")."',
  AutoUpdateModules: {
    18: '        Debug.Print "[AutoUpdateModules] Sunucuya baglanilamadi. HTTP " & http.Status',
    76: '    Debug.Print "[AutoUpdateModules] " & count & " modul sunucudan alindi. Moduller sayfasina yazildi."',
    81: '    Debug.Print "[AutoUpdateModules] Baglanti hatasi: " & Err.Description',
  },
  getLicense:
    '        Debug.Print "[getLicense] Baglanti hatasi: " & Err.Description',
  ImportRegistrySettings:
    '    Debug.Print "[ImportRegistrySettings] ImportRegistrySettings hatasi: " & Err.Description',
  RunMacroInWorkbook:
    '    If Err.Number <> 0 Then Debug.Print "[RunMacroInWorkbook] Hata: " & Err.Description',
  SendDailyEmailReport:
    '    Debug.Print "[SendDailyEmailReport] Rapor hatasi: " & Err.Description',
  ConnectToSqlServer:
    '    Debug.Print "[ConnectToSqlServer] SQL Hatasi: " & Err.Description',
};

function applyExactFixes(methodName, code) {
  const fix = EXACT_LINE_FIXES[methodName];
  if (!fix) return code;

  const lines = code.split(/\r?\n/);

  if (typeof fix === "string") {
    for (let i = 0; i < lines.length; i++) {
      if (/Debug\.Print/i.test(lines[i]) && /vbInformation|vbExclamation|procName|modName|http\.Status|Err\.Description|\(r-1\)/i.test(lines[i])) {
        lines[i] = fix;
        break;
      }
    }
  } else if (Array.isArray(fix)) {
    let fi = 0;
    for (let i = 0; i < lines.length && fi < fix.length; i++) {
      if (/Debug\.Print/i.test(lines[i]) && /modName/i.test(lines[i])) {
        lines[i] = fix[fi++];
      }
    }
  } else if (typeof fix === "object") {
    for (const [lineNum, lineText] of Object.entries(fix)) {
      lines[Number(lineNum) - 1] = lineText;
    }
  }

  return lines.join("\n");
}

function fixRuntime(methodName, code) {
  let c = code;
  const tag = `[${methodName}]`;

  c = applyExactFixes(methodName, c);

  // Debug.Print sonrasinda MsgBox parametreleri kalmis
  c = c.replace(
    /Debug\.Print\s+([^,\n]+),\s*vb\w+(\s*,\s*"[^"]*")?\s*$/gim,
    "Debug.Print $1",
  );

  // Derleme hatasi: IIf(exists"), vbInformation
  c = c.replace(
    /Debug\.Print\s+"[^"]*IIf\([^"]*"\),\s*vb\w+[^\\n]*/gi,
    (line) => {
      if (methodName === "FileExists") {
        return '    Debug.Print "[FileExists] " & fp & Chr(10) & IIf(exists, "Evet", "Hayir")';
      }
      return line.replace(/,\s*vb\w+.*$/, "");
    },
  );

  // If Not fso.FileExists(fp) Then Debug.Print "... fp ..." — dosya yok mesaji
  c = c.replace(
    /If\s+Not\s+fso\.FileExists\(fp\)\s+Then\s+Debug\.Print\s+"(\[[^\]]+\])\s*Dosya bulunamadı\."/gi,
    'If Not fso.FileExists(fp) Then Debug.Print "$1 Dosya bulunamadi: " & fp : Set DynamicFunc = Nothing : Exit Function',
  );

  // AutoUpdateModules: 686 modul yazmak yerine ilk 80
  if (methodName === "AutoUpdateModules") {
    c = c.replace(
      /If gRow >= 500 Then Exit For/gi,
      "If gRow >= 500 Then Exit For",
    );
    if (!c.includes("If row > 83 Then Exit Do")) {
      c = c.replace(
        /(\s+row = row \+ 1\s*\r?\n\s+pos = mPos \+ 14)/,
        "$1\n        If row > 83 Then Exit Do",
      );
    }
    c = c.replace(/\s*ws\.Activate\s*\r?\n/, "\n");
  }

  // Set DynamicFunc = Nothing eksik olanlar
  if (/Public Function DynamicFunc/i.test(c) && !/Set DynamicFunc\s*=\s*Nothing/i.test(c)) {
    c = c.replace(
      /(End Function\s*)$/i,
      "    Set DynamicFunc = Nothing\nEnd Function\n",
    );
  }

  // GenerateRandomToken returns Set DynamicFunc = s — OK

  // AutoFit yavasligi
  c = c.replace(/ws\.Columns\.AutoFit/g, "ws.UsedRange.Columns.AutoFit");
  c = c.replace(/ws\.Columns\("A:C"\)\.AutoFit/g, 'ws.Range("A1:C" & row - 1).Columns.AutoFit');

  // Stub moduller
  if (/WinAPI Declare; bkz dll-module-proposals/i.test(c)) {
    c = `Public Function DynamicFunc(targetWb As Workbook, param As Variant) As Object
    Debug.Print "${tag} Stub — henuz WinAPI implementasyonu yok. param=" & CStr(param)
    Set DynamicFunc = Nothing
End Function
`;
  }

  return c;
}

function hasCompileRisk(code) {
  return (
    /vbInformation|vbExclamation|vbCritical|vbYesNo/i.test(code) ||
    /Debug\.Print[^\\n]*\),\s*vb/i.test(code) ||
    /Debug\.Print[^\\n]*IIf\([^)]*"\)/i.test(code)
  );
}

const rows = await sql`SELECT method_name, code FROM modules ORDER BY method_name`;
let updated = 0;
const touched = [];

for (const row of rows) {
  const fixed = fixRuntime(row.method_name, row.code);
  if (fixed === row.code) continue;

  const stillRisk = hasCompileRisk(fixed);
  if (stillRisk) {
    console.warn("UYARI hala risk:", row.method_name);
  }

  touched.push(row.method_name);
  if (!dryRun) {
    await sql`
      UPDATE modules SET code = ${fixed}, updated_at = NOW()
      WHERE method_name = ${row.method_name}
    `;
  }
  updated++;
}

const outDir = path.join(process.cwd(), "data", "audit");
fs.mkdirSync(outDir, { recursive: true });
fs.writeFileSync(
  path.join(outDir, "runtime-fix-log.json"),
  JSON.stringify({ updated: touched, count: updated }, null, 2),
);

console.log(dryRun ? "[dry-run] " : "", `Duzeltilen modul: ${updated}`);
touched.slice(0, 40).forEach((n) => console.log(" -", n));
if (touched.length > 40) console.log(` ... +${touched.length - 40} daha`);
