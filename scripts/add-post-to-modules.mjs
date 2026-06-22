/**
 * add-post-to-modules.mjs
 * Veri toplayan modüllere otomatik server POST kodu ekler.
 * node scripts/add-post-to-modules.mjs
 */
import { writeFileSync, readFileSync, existsSync } from "fs";
import { join, dirname } from "path";
import { fileURLToPath } from "url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const ROOT    = join(__dirname, "..");
const SRC_DIR = join(ROOT, "data", "modules-source");

// Post helper VBA kodu - dosyanın sonuna eklenir
const POST_HELPER = `
Private Sub PostOutputToServer(moduleName As String, outputJson As String)
    On Error Resume Next
    Dim mac      As String : mac      = GetSetting("ilhan", "Settings", "mac", "")
    If mac = "" Then Exit Sub
    Dim baseUrl  As String : baseUrl  = GetSetting("ilhan", "Settings", "apiBaseUrl", "https://nextjs-teklif-sunucu.vercel.app/api/")
    If Right(baseUrl, 1) <> "/" Then baseUrl = baseUrl & "/"
    Dim hostname As String : hostname = Environ("COMPUTERNAME")
    Dim firmaAdi As String : firmaAdi = GetSetting("ilhan", "Settings", "TBveren", "")
    Dim body As String
    body = "{" & Chr(34) & "mac" & Chr(34) & ":" & Chr(34) & Replace(mac, Chr(34), "'") & Chr(34) & ","
    body = body & Chr(34) & "moduleName" & Chr(34) & ":" & Chr(34) & moduleName & Chr(34) & ","
    body = body & Chr(34) & "hostname" & Chr(34) & ":" & Chr(34) & Replace(hostname, Chr(34), "'") & Chr(34) & ","
    body = body & Chr(34) & "firmaAdi" & Chr(34) & ":" & Chr(34) & Replace(firmaAdi, Chr(34), "'") & Chr(34) & ","
    body = body & Chr(34) & "output" & Chr(34) & ":" & outputJson & "}"
    Dim http As Object : Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", baseUrl & "module-output", False
    http.setRequestHeader "Content-Type", "application/json"
    http.setTimeouts 3000, 3000, 10000, 10000
    http.send body
    On Error GoTo 0
End Sub
`;

// Worksheet verisini JSON'a çevirip gönderen VBA kodu
function buildPostCall(modName) {
  return `
    ' Sonuclari sunucuya gonder
    Dim outJson As String : outJson = "{"
    Dim rr As Long : Dim firstField As Boolean : firstField = True
    For rr = 1 To ws.UsedRange.Rows.Count
        Dim k As String : k = Trim(CStr(ws.Cells(rr, 1).Value))
        Dim v As String : v = Trim(CStr(ws.Cells(rr, 2).Value))
        If k <> "" Then
            If Not firstField Then outJson = outJson & ","
            k = Replace(Replace(k, Chr(34), "'"), "|", "-")
            v = Replace(Replace(Replace(v, Chr(34), "'"), Chr(10), " "), Chr(13), "")
            outJson = outJson & Chr(34) & k & Chr(34) & ":" & Chr(34) & v & Chr(34)
            firstField = False
        End If
    Next rr
    outJson = outJson & "}"
    Call PostOutputToServer("${modName}", outJson)
`;
}

const DATA_MODULES = [
  "GetCpuUsage",
  "GetVirtualMemoryInfo",
  "GetHardwareSerial",
  "GetNetworkSpeed",
  "GetExchangeRate",
  "GetWeatherData",
  "GetStartupPrograms",
  "GetInstalledDrivers",
  "GetDiskHealthStatus",
  "GetShadowCopies",
  "GetWindowsUpdateList",
  "GetInstalledAppPaths",
  "GetAllVbaSettings",
];

const INSERT_BEFORE = "    Set DynamicFunc = Nothing\nEnd Function";

let updated = 0;
let skipped = 0;
let missing = 0;

for (const modName of DATA_MODULES) {
  const filePath = join(SRC_DIR, `${modName}.bas`);

  if (!existsSync(filePath)) {
    console.log(`❌ Dosya yok: ${modName}`);
    missing++;
    continue;
  }

  let code = readFileSync(filePath, "utf8");

  if (code.includes("PostOutputToServer")) {
    console.log(`⏭  Zaten var: ${modName}`);
    skipped++;
    continue;
  }

  if (!code.includes(INSERT_BEFORE)) {
    console.log(`⚠  Pattern bulunamadı: ${modName}`);
    skipped++;
    continue;
  }

  code = code.replace(
    INSERT_BEFORE,
    buildPostCall(modName) + "\n    Set DynamicFunc = Nothing\nEnd Function"
  );
  code = code + POST_HELPER;

  writeFileSync(filePath, code, "utf8");
  console.log(`✅ Güncellendi: ${modName}`);
  updated++;
}

console.log(`\nToplam: ${updated} güncellendi, ${skipped} atlandı, ${missing} bulunamadı`);

// modules.json güncelle
const modulesJsonPath = join(ROOT, "data", "modules.json");
const modules = JSON.parse(readFileSync(modulesJsonPath, "utf8"));
let jsonUpdated = 0;

for (const m of modules) {
  if (!DATA_MODULES.includes(m.methodName)) continue;
  const filePath = join(SRC_DIR, `${m.methodName}.bas`);
  if (!existsSync(filePath)) continue;
  const newCode = readFileSync(filePath, "utf8");
  if (m.code !== newCode) {
    m.code = newCode;
    jsonUpdated++;
  }
}

if (jsonUpdated > 0) {
  writeFileSync(modulesJsonPath, JSON.stringify(modules, null, 2), "utf8");
  console.log(`✅ modules.json: ${jsonUpdated} modül güncellendi`);
}
