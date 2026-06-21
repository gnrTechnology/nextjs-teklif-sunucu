import fs from "fs";
import path from "path";
import { neon } from "@neondatabase/serverless";

function loadEnvLocal() {
  const envPath = path.join(process.cwd(), ".env.local");
  if (!fs.existsSync(envPath)) {
    throw new Error(".env.local bulunamadı. DATABASE_URL tanımlayın.");
  }
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

// Tablo oluştur
await sql`
  CREATE TABLE IF NOT EXISTS licenses (
    mac_adresi VARCHAR(17) PRIMARY KEY,
    ip_adresi  VARCHAR(45),
    firma_adi  VARCHAR(255),
    user_adi   VARCHAR(255),
    dosya_adi  VARCHAR(255),
    license    VARCHAR(50) NOT NULL DEFAULT 'true',
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
  )
`;

// Fazla kolonları temizle (eski şemadan kalma)
const dropCols = [
  "proje_adi", "proje_kisa_adresi",
  "teklif_para_birimi_usd", "teklif_para_birimi_euro", "teklif_para_birimi_genel",
  "genel_gider", "kar", "m31_degeri", "veritabani_teklif",
];
for (const col of dropCols) {
  await sql`ALTER TABLE licenses DROP COLUMN IF EXISTS ${sql.unsafe(col)}`;
}

// user_adi yoksa ekle
await sql`ALTER TABLE licenses ADD COLUMN IF NOT EXISTS user_adi VARCHAR(255)`;
await sql`CREATE INDEX IF NOT EXISTS idx_licenses_firma_adi ON licenses (firma_adi)`;

// Denetim logu tablosu
await sql`
  CREATE TABLE IF NOT EXISTS license_logs (
    id         BIGSERIAL PRIMARY KEY,
    mac_adresi VARCHAR(17),
    firma_adi  VARCHAR(255),
    user_adi   VARCHAR(255),
    dosya_adi  VARCHAR(255),
    ip_adresi  VARCHAR(45),
    event_type VARCHAR(50) NOT NULL DEFAULT 'register',
    details    TEXT,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
  )
`;
await sql`CREATE INDEX IF NOT EXISTS idx_license_logs_mac ON license_logs (mac_adresi)`;
await sql`CREATE INDEX IF NOT EXISTS idx_license_logs_created ON license_logs (created_at DESC)`;

console.log("Neon licenses tablosu hazır.");
console.log("Kolonlar: mac_adresi, ip_adresi, firma_adi, user_adi, dosya_adi, license, created_at, updated_at");
