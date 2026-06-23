/**
 * FolderWatchPollHelpers.bas → InstallCommandQueue.bas içinde FwPollHelpersCode() üretir.
 * Her satır ayrı s = s & "..." & vbCrLf ile eklenir (tırnak/vbCrLf hatası önlenir).
 */
const fs = require("fs");
const path = require("path");

const root = path.join(__dirname, "..");
const helpersPath = path.join(root, "data", "modules-source", "FolderWatchPollHelpers.bas");
const installPath = path.join(root, "data", "modules-source", "InstallCommandQueue.bas");

const helpers = fs.readFileSync(helpersPath, "utf8").trim().replace(/\r?\n/g, "\n");
const lines = helpers.split("\n");

function vbaEscape(line) {
  return line.replace(/"/g, '""');
}

let fn = "Private Function FwPollHelpersCode() As String\r\n";
fn += "    Dim s As String\r\n";
for (const line of lines) {
  fn += `    s = s & "${vbaEscape(line)}" & vbCrLf\r\n`;
}
fn += "    FwPollHelpersCode = s\r\n";
fn += "End Function\r\n";

let install = fs.readFileSync(installPath, "utf8");
const startMarker = "Private Function FwPollHelpersCode() As String";
const endMarker = "Private Function GetPollModuleCode() As String";

const startIdx = install.indexOf(startMarker);
const endIdx = install.indexOf(endMarker);
if (startIdx < 0 || endIdx < 0 || endIdx <= startIdx) {
  console.error("InstallCommandQueue.bas icinde FwPollHelpersCode bulunamadi.");
  process.exit(1);
}

install =
  install.slice(0, startIdx) +
  fn.replace(/\n/g, "\r\n") +
  "\r\n" +
  install.slice(endIdx);

fs.writeFileSync(installPath, install);
console.log(`FwPollHelpersCode guncellendi (${lines.length} satir).`);
