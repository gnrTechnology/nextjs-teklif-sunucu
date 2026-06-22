const fs = require("fs");
const path = require("path");

const root = path.join(__dirname, "..");
const srcDir = path.join(root, "data", "modules-source");
const modulesPath = path.join(root, "data", "modules.json");

const WMI_CORRECT = "winmgmts:\\\\." + "\\root\\cimv2";

function fixWmiPaths(code) {
  let c = code;
  c = c.replace(/winmgmts:(\\+\.)rootcimv2/gi, `winmgmts:$1\\root\\cimv2`);
  c = c.replace(/winmgmts:(\\+\.)ootcimv2/gi, `winmgmts:$1\\root\\cimv2`);
  c = c.replace(/winmgmts:(\\+\.)\s+ootcimv2/gi, `winmgmts:$1\\root\\cimv2`);
  return c;
}

let count = 0;

for (const file of fs.readdirSync(srcDir)) {
  if (!file.endsWith(".bas")) continue;
  const fp = path.join(srcDir, file);
  const raw = fs.readFileSync(fp, "utf8");
  const fixed = fixWmiPaths(raw);
  if (fixed !== raw) {
    fs.writeFileSync(fp, fixed, "utf8");
    console.log("bas:", file);
    count++;
  }
}

const modules = JSON.parse(fs.readFileSync(modulesPath, "utf8"));
for (const m of modules) {
  if (!m.code) continue;
  const fixed = fixWmiPaths(m.code);
  if (fixed !== m.code) {
    m.code = fixed;
    const fromBas = path.join(srcDir, `${m.methodName}.bas`);
    if (fs.existsSync(fromBas)) {
      m.code = fs.readFileSync(fromBas, "utf8");
    }
    console.log("json:", m.methodName);
    count++;
  }
}

for (const m of modules) {
  const fromBas = path.join(srcDir, `${m.methodName}.bas`);
  if (fs.existsSync(fromBas)) {
    const basCode = fs.readFileSync(fromBas, "utf8");
    if (basCode.includes("root\\cimv2") && m.code.includes("rootcimv2")) {
      m.code = basCode;
      console.log("sync from bas:", m.methodName);
      count++;
    }
  }
}

fs.writeFileSync(modulesPath, JSON.stringify(modules, null, 2), "utf8");
console.log("done, changes:", count);
