#!/usr/bin/env node
// Cursor stop hook: ajan turunu bitirince degisiklik varsa otomatik commit atar.

const { execSync } = require("child_process");

function run(cmd) {
  try {
    return execSync(cmd, { encoding: "utf8", stdio: ["pipe", "pipe", "pipe"] }).trim();
  } catch (e) {
    return "";
  }
}

const status = run("git status --porcelain");

if (!status) {
  // Degisiklik yok, bir sey yapma
  process.stdout.write("{}");
  process.exit(0);
}

// Degisen dosyalarin ozet listesi (ilk 3 dosya)
const lines = status.split("\n").filter(Boolean);
const fileList = lines
  .slice(0, 3)
  .map((l) => l.trim().replace(/^\S+\s+/, ""))
  .join(", ");
const more = lines.length > 3 ? ` +${lines.length - 3} daha` : "";

const now = new Date().toLocaleString("tr-TR", {
  day: "2-digit",
  month: "2-digit",
  hour: "2-digit",
  minute: "2-digit",
});

const msg = `auto(${now}): ${fileList}${more}`;

run("git add -A");
run(`git commit -m "${msg.replace(/"/g, "'")}"`);

process.stdout.write("{}");
process.exit(0);
