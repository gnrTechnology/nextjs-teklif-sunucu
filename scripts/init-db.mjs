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
    if (!process.env[key]) {
      process.env[key] = value;
    }
  }
}

loadEnvLocal();

const databaseUrl = process.env.DATABASE_URL;
if (!databaseUrl) {
  throw new Error("DATABASE_URL tanımlı değil.");
}

const sql = neon(databaseUrl);

await sql`
  CREATE TABLE IF NOT EXISTS licenses (
    mac_adresi VARCHAR(17) PRIMARY KEY,
    ip_adresi VARCHAR(45),
    firma_adi VARCHAR(255),
    user_adi VARCHAR(255),
    dosya_adi VARCHAR(255),
    proje_adi VARCHAR(255),
    proje_kisa_adresi TEXT,
    teklif_para_birimi_usd VARCHAR(50),
    teklif_para_birimi_euro VARCHAR(50),
    teklif_para_birimi_genel VARCHAR(50),
    genel_gider VARCHAR(50),
    kar VARCHAR(50),
    m31_degeri VARCHAR(50),
    veritabani_teklif JSONB,
    license VARCHAR(50) NOT NULL DEFAULT 'true',
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
  )
`;

await sql`ALTER TABLE licenses ADD COLUMN IF NOT EXISTS user_adi VARCHAR(255)`;
await sql`CREATE INDEX IF NOT EXISTS idx_licenses_firma_adi ON licenses (firma_adi)`;

console.log("Neon licenses tablosu hazır.");
