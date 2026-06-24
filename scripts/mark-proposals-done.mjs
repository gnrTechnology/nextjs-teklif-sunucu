/**
 * module-proposals.md ve dll-module-proposals.md icindeki modulleri isaretler
 */
import fs from "fs";
import path from "path";

const names = new Set(
  fs
    .readdirSync(path.join(process.cwd(), "data", "modules-new"))
    .filter((f) => f.endsWith(".bas"))
    .map((f) => f.replace(/\.bas$/i, "")),
);

function markFile(filePath) {
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

const a = markFile(path.join(process.cwd(), "data", "module-proposals.md"));
const b = markFile(path.join(process.cwd(), "data", "dll-module-proposals.md"));
console.log("marked:", { "module-proposals.md": a, "dll-module-proposals.md": b });
