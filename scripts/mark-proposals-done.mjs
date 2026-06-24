/**
 * module-proposals.md ve dll-module-proposals.md icindeki modulleri isaretler.
 * Kaynak: data/modules-staging/*.bas veya data/modules-meta.json
 */
import fs from "fs";
import path from "path";

const ROOT = process.cwd();
const STAGING = path.join(ROOT, "data", "modules-staging");
const META_PATH = path.join(ROOT, "data", "modules-meta.json");

function collectNames() {
  const names = new Set();
  if (fs.existsSync(STAGING)) {
    for (const f of fs.readdirSync(STAGING)) {
      if (f.endsWith(".bas")) names.add(f.replace(/\.bas$/i, ""));
    }
  }
  if (names.size === 0 && fs.existsSync(META_PATH)) {
    for (const key of Object.keys(JSON.parse(fs.readFileSync(META_PATH, "utf8")))) {
      names.add(key);
    }
  }
  return names;
}

function markFile(filePath, names) {
  if (!fs.existsSync(filePath)) return 0;
  let md = fs.readFileSync(filePath, "utf8");
  let updated = 0;
  for (const name of names) {
    const re = new RegExp(
      `(\\|\\s*(?:D\\d+|S\\d+|H\\d+|\\d+[a-z]?)\\s*\\|\\s*${name}\\s*\\|(?:[^|\\n]*\\|)*\\s*(?:🔒\\s*)?(?:⚠️\\s*)?)⬜`,
      "g",
    );
    const next = md.replace(re, "$1✅");
    if (next !== md) {
      updated++;
      md = next;
    }
  }
  fs.writeFileSync(filePath, md, "utf8");
  return updated;
}

const names = collectNames();
const a = markFile(path.join(ROOT, "data", "module-proposals.md"), names);
const b = markFile(path.join(ROOT, "data", "dll-module-proposals.md"), names);
console.log("marked:", { "module-proposals.md": a, "dll-module-proposals.md": b, names: names.size });
