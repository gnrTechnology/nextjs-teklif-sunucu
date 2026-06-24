/**
 * olevba ciktisini modul dosyalarina ayirir.
 * node scripts/xlam-split-olevba.mjs
 */
import fs from "fs";
import path from "path";

const src = path.join("data", "xlam-backup", "olevba-all.txt");
const outDir = path.join("data", "xlam-backup", "modules");

const text = fs.readFileSync(src, "utf8");
fs.mkdirSync(outDir, { recursive: true });

const parts = text.split(/^VBA MACRO /m).filter(Boolean);
const index = [];

for (const part of parts) {
  const nl = part.indexOf("\n");
  if (nl < 0) continue;
  const header = part.slice(0, nl).trim();
  const nameMatch = header.match(/^([\w.]+\.(?:bas|cls|frm))/i);
  if (!nameMatch) continue;
  const rawName = nameMatch[1];
  const safeName = rawName.replace(/[^\w.-]/g, "_");

  const bodyStart = part.indexOf("- - -");
  let code = part;
  if (bodyStart >= 0) {
    const afterDashes = part.indexOf("\n", bodyStart);
    const endDash = part.indexOf("\n-------------------------------------------------------------------------------", afterDashes + 1);
    code = endDash > afterDashes ? part.slice(afterDashes + 1, endDash) : part.slice(afterDashes + 1);
  }

  code = code.trim();
  const lines = code.split(/\r?\n/).length;
  const ext = rawName.includes(".cls") || rawName.endsWith(".cls") ? ".cls" : rawName.includes(".frm") ? ".frm" : ".bas";
  const fileName = safeName.replace(/\.(bas|cls|frm)$/i, "") + ext;
  const filePath = path.join(outDir, fileName);
  fs.writeFileSync(filePath, code, "utf8");
  index.push({ name: rawName, file: fileName, lines });
}

index.sort((a, b) => b.lines - a.lines);
fs.writeFileSync(
  path.join("data", "xlam-backup", "index.json"),
  JSON.stringify({ extractedAt: new Date().toISOString(), count: index.length, modules: index }, null, 2),
);

console.log("Modul:", index.length, "->", outDir);
