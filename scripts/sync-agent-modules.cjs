const fs = require("fs");
const path = require("path");

const root = path.join(__dirname, "..");

function readBas(name) {
  return fs.readFileSync(path.join(root, "data", "modules-source", name), "utf8");
}

const modules = JSON.parse(
  fs.readFileSync(path.join(root, "data", "modules.json"), "utf8"),
);

const hbIdx = modules.findIndex((m) => m.methodName === "HeartbeatPing");
if (hbIdx >= 0) {
  modules[hbIdx].code = readBas("HeartbeatPing.bas");
  modules[hbIdx].description =
    "Anlik ping + TeklifAgent DLL ile arka plan heartbeat ve uzak komut kuyrugu";
}

const installEntry = {
  methodName: "InstallTeklifAgent",
  description: "TeklifAgent COM DLL + exe indirir ve regasm ile kaydeder",
  category: "zamanlanmis",
  active: true,
  code: readBas("InstallTeklifAgent.bas"),
};

const installIdx = modules.findIndex((m) => m.methodName === "InstallTeklifAgent");
if (installIdx >= 0) modules[installIdx] = installEntry;
else modules.splice(hbIdx + 1, 0, installEntry);

fs.writeFileSync(
  path.join(root, "data", "modules.json"),
  JSON.stringify(modules, null, 2),
  "utf8",
);
console.log("modules.json guncellendi");
