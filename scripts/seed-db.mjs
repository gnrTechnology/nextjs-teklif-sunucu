import fs from "fs";
import path from "path";
import { neon } from "@neondatabase/serverless";

function loadEnvLocal() {
  const envPath = path.join(process.cwd(), ".env.local");
  if (!fs.existsSync(envPath)) throw new Error(".env.local bulunamadı.");
  for (const line of fs.readFileSync(envPath, "utf-8").split(/\r?\n/)) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("#")) continue;
    const eq = trimmed.indexOf("=");
    if (eq === -1) continue;
    const key = trimmed.slice(0, eq).trim();
    const value = trimmed.slice(eq + 1).trim();
    if (!process.env[key]) process.env[key] = value;
  }
}

loadEnvLocal();
const sql = neon(process.env.DATABASE_URL);

// 1. Test kaydını sil
await sql`DELETE FROM licenses WHERE mac_adresi = 'AA:BB:CC:DD:EE:FF'`;
console.log("Test kaydı silindi: AA:BB:CC:DD:EE:FF");

// 2. licenses.json'daki gerçek EPRON kaydını ekle
const inserted = await sql`
  INSERT INTO licenses (mac_adresi, ip_adresi, firma_adi, dosya_adi, license, created_at, updated_at)
  VALUES (
    '04:EC:D8:AE:C0:4A',
    '192.168.1.159',
    'EPRON',
    'teklif.xlam',
    'true',
    '2026-06-20T13:41:54.262Z',
    '2026-06-20T13:45:01.974Z'
  )
  ON CONFLICT (mac_adresi) DO NOTHING
  RETURNING mac_adresi
`;
console.log(inserted.length > 0 ? `Eklendi: ${inserted[0].mac_adresi}` : "EPRON kaydı zaten vardı, dokunulmadı.");

// 3. Mevcut tüm kayıtlar
const all = await sql`SELECT mac_adresi, firma_adi, user_adi, license, updated_at FROM licenses ORDER BY updated_at DESC`;
console.log("\nMevcut kayıtlar:");
for (const r of all) {
  console.log(` • ${r.mac_adresi} | ${r.firma_adi ?? "—"} | ${r.license} | ${r.updated_at}`);
}
