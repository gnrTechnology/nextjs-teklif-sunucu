const fs = require("fs");
const path = require("path");

const root = path.join(__dirname, "..");
const srcDir = path.join(root, "data", "modules-source");
const modulesPath = path.join(root, "data", "modules.json");

const getMacBlock = `    If mac = "" Then mac = GetMacFromWmi()
    If mac = "" Then Exit Sub`;

const getMacFn = `
Private Function GetMacFromWmi() As String
    On Error Resume Next
    Dim wmi As Object, col As Object, o As Object
    Set wmi = GetObject("winmgmts:\\\\.` + `\\root\\cimv2")
    Set col = wmi.ExecQuery("SELECT MACAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=True")
    For Each o In col
        If Not IsNull(o.MACAddress) And o.MACAddress <> "" Then
            GetMacFromWmi = o.MACAddress
            Exit Function
        End If
    Next
End Function
`;

const updated = [];

for (const file of fs.readdirSync(srcDir)) {
  if (!file.endsWith(".bas")) continue;
  const fp = path.join(srcDir, file);
  let code = fs.readFileSync(fp, "utf8");
  if (!code.includes("PostOutputToServer")) continue;

  let changed = false;
  if (code.includes('baseUrl & "module-output"')) {
    code = code.replace(/baseUrl & "module-output"/g, 'baseUrl & "module-output/"');
    changed = true;
  }
  if (
    code.includes('GetSetting("ilhan", "Settings", "mac", "")') &&
    !code.includes("GetMacFromWmi()")
  ) {
    code = code.replace(
      /Dim mac\s+As String\s*:\s*mac\s*=\s*GetSetting\("ilhan", "Settings", "mac", ""\)\r?\n\s*If mac = "" Then Exit Sub/,
      `Dim mac      As String : mac      = GetSetting("ilhan", "Settings", "mac", "")\r\n${getMacBlock}`,
    );
    changed = true;
  }
  if (!code.includes("Function GetMacFromWmi")) {
    code = code.replace(/\r?\n\s*$/, "") + getMacFn + "\r\n";
    changed = true;
  }
  if (changed) {
    fs.writeFileSync(fp, code, "utf8");
    updated.push(path.basename(file, ".bas"));
    console.log("fixed:", file);
  }
}

if (updated.length) {
  const modules = JSON.parse(fs.readFileSync(modulesPath, "utf8"));
  for (const name of updated) {
    const code = fs.readFileSync(path.join(srcDir, `${name}.bas`), "utf8");
    const idx = modules.findIndex((m) => m.methodName === name);
    if (idx >= 0) modules[idx].code = code;
  }
  fs.writeFileSync(modulesPath, JSON.stringify(modules, null, 2), "utf8");
  console.log("modules.json guncellendi:", updated.join(", "));
}
