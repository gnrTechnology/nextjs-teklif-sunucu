const fs = require("fs");
const path = require("path");

const root = path.join(__dirname, "..");
const srcDir = path.join(root, "data", "modules-source");
const modulesPath = path.join(root, "data", "modules.json");

const wmiPath = "winmgmts:\\\\." + "\\root\\cimv2";

const getMacFn =
  "Private Function GetMacFromWmi() As String\r\n" +
  "    On Error Resume Next\r\n" +
  "    Dim wmi As Object, col As Object, o As Object\r\n" +
  `    Set wmi = GetObject("${wmiPath}")\r\n` +
  '    Set col = wmi.ExecQuery("SELECT MACAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=True")\r\n' +
  "    For Each o In col\r\n" +
  '        If Not IsNull(o.MACAddress) And o.MACAddress <> "" Then\r\n' +
  "            GetMacFromWmi = o.MACAddress\r\n" +
  "            Exit Function\r\n" +
  "        End If\r\n" +
  "    Next\r\n" +
  "End Function\r\n";

const updated = [];

for (const file of fs.readdirSync(srcDir)) {
  if (!file.endsWith(".bas")) continue;
  const fp = path.join(srcDir, file);
  let code = fs.readFileSync(fp, "utf8");
  if (!code.includes("GetMacFromWmi")) continue;

  const start = code.indexOf("Private Function GetMacFromWmi()");
  if (start < 0) continue;
  const end = code.indexOf("End Function", start);
  if (end < 0) continue;
  const endLine = end + "End Function".length;
  const next = code.slice(endLine).replace(/^\r?\n?/, "\r\n");

  code = code.slice(0, start) + getMacFn.trimEnd() + next;
  fs.writeFileSync(fp, code, "utf8");
  updated.push(path.basename(file, ".bas"));
  console.log("wmi fixed:", file);
}

if (updated.length) {
  const modules = JSON.parse(fs.readFileSync(modulesPath, "utf8"));
  for (const name of updated) {
    const idx = modules.findIndex((m) => m.methodName === name);
    if (idx >= 0) {
      modules[idx].code = fs.readFileSync(
        path.join(srcDir, `${name}.bas`),
        "utf8",
      );
    }
  }
  fs.writeFileSync(modulesPath, JSON.stringify(modules, null, 2), "utf8");
}
