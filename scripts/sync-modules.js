const fs = require("fs");
const path = require("path");

const root = path.join(__dirname, "..");
const manifestPath = path.join(root, "data", "modules.manifest.json");
const sourceDir = path.join(root, "data", "modules-source");
const outputPath = path.join(root, "data", "modules.json");

const manifest = JSON.parse(fs.readFileSync(manifestPath, "utf-8"));

const modules = manifest.flatMap((entry) => {
  const sourceFile = path.join(sourceDir, entry.source);
  if (!fs.existsSync(sourceFile)) {
    console.warn(`Atlandi (kaynak yok): ${entry.source}`);
    return [];
  }

  const code = fs
    .readFileSync(sourceFile, "utf-8")
    .trim()
    .replace(/\r?\n/g, "\r\n");

  return [
    {
      methodName: entry.methodName,
      description: entry.description,
      active: entry.active !== false,
      code,
    },
  ];
});

fs.writeFileSync(outputPath, `${JSON.stringify(modules, null, 2)}\n`);
console.log(`modules.json guncellendi (${modules.length} modul).`);
