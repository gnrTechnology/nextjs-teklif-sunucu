/**
 * teklif.xlam VBA yedekleme yardimcisi.
 * Gercek yedekleme VBA Developer MCP vba_backup araci ile yapilir.
 *
 * node scripts/xlam-backup.mjs
 */
import fs from "fs";
import path from "path";

const xlam =
  process.argv[2] ||
  process.env.TEKLIF_XLAM ||
  path.join(process.env.APPDATA || "", "Microsoft", "AddIns", "teklif.xlam");

const outDir = path.join(process.cwd(), "data", "xlam-backup");

if (!fs.existsSync(xlam)) {
  console.error("Dosya bulunamadi:", xlam);
  process.exit(1);
}

fs.mkdirSync(outDir, { recursive: true });

console.log("Hedef:", xlam);
console.log("Cikti:", outDir);
console.log("");
console.log("Cursor'da VBA Developer MCP:");
console.log(`  vba_backup(file_path="${xlam}", backup_dir="${outDir}")`);
console.log("");
console.log("Koruma hatasi icin:");
console.log("  Excel > Guven Merkezi > VBA proje nesne modeline erişime güven");
console.log("  VBA Editor > VBAProject Properties > Protection kaldirin");
