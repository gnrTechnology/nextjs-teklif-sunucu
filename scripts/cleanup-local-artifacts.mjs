/**
 * Yerel artefaktlari temizler (Neon DB kaynak dogrulugu).
 * node scripts/cleanup-local-artifacts.mjs
 */
import fs from "fs";
import path from "path";

const ROOT = process.cwd();
const KEEP_SOURCE = new Set([
  "zInternet-additions.bas",
  "TeklifBootstrap.bas",
  "FolderWatchPollHelpers.bas",
  "ExportRegistryOnce.bas",
  "HeartbeatPing.bas",
  "InstallCommandQueue.bas",
  "InstallTeklifAgent.bas",
]);

function rmDir(dir) {
  if (!fs.existsSync(dir)) return 0;
  let n = 0;
  for (const f of fs.readdirSync(dir)) {
    const p = path.join(dir, f);
    if (fs.statSync(p).isDirectory()) n += rmDir(p);
    else {
      fs.unlinkSync(p);
      n++;
    }
  }
  fs.rmdirSync(dir);
  return n;
}

const srcDir = path.join(ROOT, "data", "modules-source");
let removedSource = 0;
if (fs.existsSync(srcDir)) {
  for (const f of fs.readdirSync(srcDir)) {
    if (!f.endsWith(".bas") || KEEP_SOURCE.has(f)) continue;
    fs.unlinkSync(path.join(srcDir, f));
    removedSource++;
  }
}

const legacyNew = path.join(ROOT, "data", "modules-new");
const removedNew = rmDir(legacyNew);

const staging = path.join(ROOT, "data", "modules-staging");
fs.mkdirSync(staging, { recursive: true });
if (!fs.existsSync(path.join(staging, ".gitkeep"))) {
  fs.writeFileSync(path.join(staging, ".gitkeep"), "\n");
}

console.log({
  removedModulesNew: removedNew,
  removedModulesSource: removedSource,
  keptModulesSource: [...KEEP_SOURCE].filter((f) => fs.existsSync(path.join(srcDir, f))),
});
