/**
 * module-proposals.md icindeki modul adlarini ✅ yapar (modules-new + mevcut meta)
 */
import fs from "fs";
import path from "path";

const names = new Set(
  fs
    .readdirSync(path.join(process.cwd(), "data", "modules-new"))
    .filter((f) => f.endsWith(".bas"))
    .map((f) => f.replace(/\.bas$/i, "")),
);

const mdPath = path.join(process.cwd(), "data", "module-proposals.md");
let md = fs.readFileSync(mdPath, "utf8");
let updated = 0;

for (const name of names) {
  const re = new RegExp(
    `(\\|\\s*\\d+[a-z]?\\s*\\|\\s*${name}\\s*\\|[^|]*\\|\\s*)⬜`,
    "g",
  );
  const next = md.replace(re, "$1✅");
  if (next !== md) {
    updated++;
    md = next;
  }
}

fs.writeFileSync(mdPath, md, "utf8");
console.log("marked done in proposals:", updated);
