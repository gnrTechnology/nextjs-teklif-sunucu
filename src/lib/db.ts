<<<<<<< HEAD
import { neon } from "@neondatabase/serverless";
import type { LicensePostBody, LicenseRecord, VeritabaniTeklifItem } from "./types";

type LicenseRow = {
  mac_adresi: string;
  ip_adresi: string | null;
  firma_adi: string | null;
  user_adi: string | null;
  dosya_adi: string | null;
  proje_adi: string | null;
  proje_kisa_adresi: string | null;
  teklif_para_birimi_usd: string | null;
  teklif_para_birimi_euro: string | null;
  teklif_para_birimi_genel: string | null;
  genel_gider: string | null;
  kar: string | null;
  m31_degeri: string | null;
  veritabani_teklif: VeritabaniTeklifItem[] | null;
  license: string;
  created_at: string;
  updated_at: string;
};

let schemaReady: Promise<void> | null = null;

function getSql() {
  const databaseUrl = process.env.DATABASE_URL;
  if (!databaseUrl) {
    throw new Error(
      "DATABASE_URL ortam değişkeni tanımlı değil. .env.example dosyasına bakın.",
    );
  }
  return neon(databaseUrl);
}

async function ensureLicensesTable(): Promise<void> {
  if (!schemaReady) {
    schemaReady = (async () => {
      const sql = getSql();
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
    })();
  }

  await schemaReady;
}

function rowToRecord(row: LicenseRow): LicenseRecord {
  return {
    macAdresi: row.mac_adresi,
    ipAdresi: row.ip_adresi ?? undefined,
    firmaAdi: row.firma_adi ?? undefined,
    userAdi: row.user_adi ?? undefined,
    dosyaAdi: row.dosya_adi ?? undefined,
    projeAdi: row.proje_adi ?? undefined,
    projeKisaAdresi: row.proje_kisa_adresi ?? undefined,
    teklifParaBirimiUSD: row.teklif_para_birimi_usd ?? undefined,
    teklifParaBirimiEuro: row.teklif_para_birimi_euro ?? undefined,
    teklifParaBirimiGenel: row.teklif_para_birimi_genel ?? undefined,
    genelGider: row.genel_gider ?? undefined,
    kar: row.kar ?? undefined,
    m31Degeri: row.m31_degeri ?? undefined,
    veritabaniTeklif: row.veritabani_teklif ?? undefined,
    license: row.license,
    createdAt: new Date(row.created_at).toISOString(),
    updatedAt: new Date(row.updated_at).toISOString(),
  };
}

/** MAC adresini tutarlı anahtar formatına çevirir (AA:BB:CC:DD:EE:FF) */
export function normalizeMac(mac: string): string {
  const cleaned = mac.replace(/[^a-fA-F0-9]/g, "").toUpperCase();
  if (cleaned.length !== 12) {
    return mac.trim().toUpperCase();
  }
  return cleaned.match(/.{1,2}/g)!.join(":");
}

export async function getLicenseByMac(
  mac: string,
): Promise<LicenseRecord | undefined> {
  await ensureLicensesTable();
  const sql = getSql();
  const normalized = normalizeMac(mac);
  const rows = await sql`
    SELECT
      mac_adresi,
      ip_adresi,
      firma_adi,
      user_adi,
      dosya_adi,
      proje_adi,
      proje_kisa_adresi,
      teklif_para_birimi_usd,
      teklif_para_birimi_euro,
      teklif_para_birimi_genel,
      genel_gider,
      kar,
      m31_degeri,
      veritabani_teklif,
      license,
      created_at,
      updated_at
    FROM licenses
    WHERE mac_adresi = ${normalized}
    LIMIT 1
  `;

  const row = rows[0] as LicenseRow | undefined;
  return row ? rowToRecord(row) : undefined;
}

export async function upsertLicense(
  body: LicensePostBody,
): Promise<LicenseRecord> {
  await ensureLicensesTable();
  const sql = getSql();
  const normalizedMac = normalizeMac(body.macAdresi);
  const existing = await getLicenseByMac(normalizedMac);
  const now = new Date().toISOString();

  const base: LicenseRecord =
    existing ??
    ({
      macAdresi: normalizedMac,
      license: "true",
      createdAt: now,
      updatedAt: now,
    } satisfies LicenseRecord);

  const updated: LicenseRecord = {
    ...base,
    macAdresi: normalizedMac,
    ipAdresi: body.ipAdresi ?? base.ipAdresi,
    firmaAdi: body.firmaAdi ?? base.firmaAdi,
    userAdi: body.userAdi ?? base.userAdi,
    dosyaAdi: body.dosyaAdi ?? base.dosyaAdi,
    projeAdi: body.projeAdi ?? base.projeAdi,
    projeKisaAdresi: body.projeKisaAdresi ?? base.projeKisaAdresi,
    teklifParaBirimiUSD: body.teklifParaBirimiUSD ?? base.teklifParaBirimiUSD,
    teklifParaBirimiEuro:
      body.teklifParaBirimiEuro ?? base.teklifParaBirimiEuro,
    teklifParaBirimiGenel:
      body.teklifParaBirimiGenel ?? base.teklifParaBirimiGenel,
    genelGider: body.genelGider ?? base.genelGider,
    kar: body.kar ?? base.kar,
    m31Degeri: body.m31Degeri ?? base.m31Degeri,
    veritabaniTeklif: body.veritabaniTeklif ?? base.veritabaniTeklif,
    updatedAt: now,
  };

  const rows = await sql`
    INSERT INTO licenses (
      mac_adresi,
      ip_adresi,
      firma_adi,
      user_adi,
      dosya_adi,
      proje_adi,
      proje_kisa_adresi,
      teklif_para_birimi_usd,
      teklif_para_birimi_euro,
      teklif_para_birimi_genel,
      genel_gider,
      kar,
      m31_degeri,
      veritabani_teklif,
      license,
      created_at,
      updated_at
    ) VALUES (
      ${updated.macAdresi},
      ${updated.ipAdresi ?? null},
      ${updated.firmaAdi ?? null},
      ${updated.userAdi ?? null},
      ${updated.dosyaAdi ?? null},
      ${updated.projeAdi ?? null},
      ${updated.projeKisaAdresi ?? null},
      ${updated.teklifParaBirimiUSD ?? null},
      ${updated.teklifParaBirimiEuro ?? null},
      ${updated.teklifParaBirimiGenel ?? null},
      ${updated.genelGider ?? null},
      ${updated.kar ?? null},
      ${updated.m31Degeri ?? null},
      ${updated.veritabaniTeklif ?? null},
      ${updated.license},
      ${existing ? existing.createdAt : now},
      ${now}
    )
    ON CONFLICT (mac_adresi) DO UPDATE SET
      ip_adresi = EXCLUDED.ip_adresi,
      firma_adi = EXCLUDED.firma_adi,
      user_adi = EXCLUDED.user_adi,
      dosya_adi = EXCLUDED.dosya_adi,
      proje_adi = EXCLUDED.proje_adi,
      proje_kisa_adresi = EXCLUDED.proje_kisa_adresi,
      teklif_para_birimi_usd = EXCLUDED.teklif_para_birimi_usd,
      teklif_para_birimi_euro = EXCLUDED.teklif_para_birimi_euro,
      teklif_para_birimi_genel = EXCLUDED.teklif_para_birimi_genel,
      genel_gider = EXCLUDED.genel_gider,
      kar = EXCLUDED.kar,
      m31_degeri = EXCLUDED.m31_degeri,
      veritabani_teklif = EXCLUDED.veritabani_teklif,
      license = EXCLUDED.license,
      updated_at = EXCLUDED.updated_at
    RETURNING
      mac_adresi,
      ip_adresi,
      firma_adi,
      user_adi,
      dosya_adi,
      proje_adi,
      proje_kisa_adresi,
      teklif_para_birimi_usd,
      teklif_para_birimi_euro,
      teklif_para_birimi_genel,
      genel_gider,
      kar,
      m31_degeri,
      veritabani_teklif,
      license,
      created_at,
      updated_at
  `;

  return rowToRecord(rows[0] as LicenseRow);
}

export async function listLicenses(): Promise<LicenseRecord[]> {
  await ensureLicensesTable();
  const sql = getSql();
  const rows = await sql`
    SELECT
      mac_adresi,
      ip_adresi,
      firma_adi,
      user_adi,
      dosya_adi,
      proje_adi,
      proje_kisa_adresi,
      teklif_para_birimi_usd,
      teklif_para_birimi_euro,
      teklif_para_birimi_genel,
      genel_gider,
      kar,
      m31_degeri,
      veritabani_teklif,
      license,
      created_at,
      updated_at
    FROM licenses
    ORDER BY updated_at DESC
  `;

  return (rows as LicenseRow[]).map(rowToRecord);
}

export async function isLicensed(mac: string): Promise<boolean> {
  const record = await getLicenseByMac(mac);
  if (!record) return false;
  const value = record.license.toLowerCase();
  return (
    value === "true" ||
    value === "1" ||
    value === "active" ||
    value === "evet"
  );
}
=======
import { neon } from "@neondatabase/serverless";
import type { LicensePostBody, LicenseRecord, VeritabaniTeklifItem } from "./types";

type LicenseRow = {
  mac_adresi: string;
  ip_adresi: string | null;
  firma_adi: string | null;
  user_adi: string | null;
  dosya_adi: string | null;
  proje_adi: string | null;
  proje_kisa_adresi: string | null;
  teklif_para_birimi_usd: string | null;
  teklif_para_birimi_euro: string | null;
  teklif_para_birimi_genel: string | null;
  genel_gider: string | null;
  kar: string | null;
  m31_degeri: string | null;
  veritabani_teklif: VeritabaniTeklifItem[] | null;
  license: string;
  created_at: string;
  updated_at: string;
};

let schemaReady: Promise<void> | null = null;

function getSql() {
  const databaseUrl = process.env.DATABASE_URL;
  if (!databaseUrl) {
    throw new Error(
      "DATABASE_URL ortam değişkeni tanımlı değil. .env.example dosyasına bakın.",
    );
  }
  return neon(databaseUrl);
}

async function ensureLicensesTable(): Promise<void> {
  if (!schemaReady) {
    schemaReady = (async () => {
      const sql = getSql();
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
    })();
  }

  await schemaReady;
}

function rowToRecord(row: LicenseRow): LicenseRecord {
  return {
    macAdresi: row.mac_adresi,
    ipAdresi: row.ip_adresi ?? undefined,
    firmaAdi: row.firma_adi ?? undefined,
    userAdi: row.user_adi ?? undefined,
    dosyaAdi: row.dosya_adi ?? undefined,
    projeAdi: row.proje_adi ?? undefined,
    projeKisaAdresi: row.proje_kisa_adresi ?? undefined,
    teklifParaBirimiUSD: row.teklif_para_birimi_usd ?? undefined,
    teklifParaBirimiEuro: row.teklif_para_birimi_euro ?? undefined,
    teklifParaBirimiGenel: row.teklif_para_birimi_genel ?? undefined,
    genelGider: row.genel_gider ?? undefined,
    kar: row.kar ?? undefined,
    m31Degeri: row.m31_degeri ?? undefined,
    veritabaniTeklif: row.veritabani_teklif ?? undefined,
    license: row.license,
    createdAt: new Date(row.created_at).toISOString(),
    updatedAt: new Date(row.updated_at).toISOString(),
  };
}

/** MAC adresini tutarlı anahtar formatına çevirir (AA:BB:CC:DD:EE:FF) */
export function normalizeMac(mac: string): string {
  const cleaned = mac.replace(/[^a-fA-F0-9]/g, "").toUpperCase();
  if (cleaned.length !== 12) {
    return mac.trim().toUpperCase();
  }
  return cleaned.match(/.{1,2}/g)!.join(":");
}

export async function getLicenseByMac(
  mac: string,
): Promise<LicenseRecord | undefined> {
  await ensureLicensesTable();
  const sql = getSql();
  const normalized = normalizeMac(mac);
  const rows = await sql`
    SELECT
      mac_adresi,
      ip_adresi,
      firma_adi,
      user_adi,
      dosya_adi,
      proje_adi,
      proje_kisa_adresi,
      teklif_para_birimi_usd,
      teklif_para_birimi_euro,
      teklif_para_birimi_genel,
      genel_gider,
      kar,
      m31_degeri,
      veritabani_teklif,
      license,
      created_at,
      updated_at
    FROM licenses
    WHERE mac_adresi = ${normalized}
    LIMIT 1
  `;

  const row = rows[0] as LicenseRow | undefined;
  return row ? rowToRecord(row) : undefined;
}

export async function upsertLicense(
  body: LicensePostBody,
): Promise<LicenseRecord> {
  await ensureLicensesTable();
  const sql = getSql();
  const normalizedMac = normalizeMac(body.macAdresi);
  const existing = await getLicenseByMac(normalizedMac);
  const now = new Date().toISOString();

  const base: LicenseRecord =
    existing ??
    ({
      macAdresi: normalizedMac,
      license: "true",
      createdAt: now,
      updatedAt: now,
    } satisfies LicenseRecord);

  const updated: LicenseRecord = {
    ...base,
    macAdresi: normalizedMac,
    ipAdresi: body.ipAdresi ?? base.ipAdresi,
    firmaAdi: body.firmaAdi ?? base.firmaAdi,
    userAdi: body.userAdi ?? base.userAdi,
    dosyaAdi: body.dosyaAdi ?? base.dosyaAdi,
    projeAdi: body.projeAdi ?? base.projeAdi,
    projeKisaAdresi: body.projeKisaAdresi ?? base.projeKisaAdresi,
    teklifParaBirimiUSD: body.teklifParaBirimiUSD ?? base.teklifParaBirimiUSD,
    teklifParaBirimiEuro:
      body.teklifParaBirimiEuro ?? base.teklifParaBirimiEuro,
    teklifParaBirimiGenel:
      body.teklifParaBirimiGenel ?? base.teklifParaBirimiGenel,
    genelGider: body.genelGider ?? base.genelGider,
    kar: body.kar ?? base.kar,
    m31Degeri: body.m31Degeri ?? base.m31Degeri,
    veritabaniTeklif: body.veritabaniTeklif ?? base.veritabaniTeklif,
    updatedAt: now,
  };

  const rows = await sql`
    INSERT INTO licenses (
      mac_adresi,
      ip_adresi,
      firma_adi,
      user_adi,
      dosya_adi,
      proje_adi,
      proje_kisa_adresi,
      teklif_para_birimi_usd,
      teklif_para_birimi_euro,
      teklif_para_birimi_genel,
      genel_gider,
      kar,
      m31_degeri,
      veritabani_teklif,
      license,
      created_at,
      updated_at
    ) VALUES (
      ${updated.macAdresi},
      ${updated.ipAdresi ?? null},
      ${updated.firmaAdi ?? null},
      ${updated.userAdi ?? null},
      ${updated.dosyaAdi ?? null},
      ${updated.projeAdi ?? null},
      ${updated.projeKisaAdresi ?? null},
      ${updated.teklifParaBirimiUSD ?? null},
      ${updated.teklifParaBirimiEuro ?? null},
      ${updated.teklifParaBirimiGenel ?? null},
      ${updated.genelGider ?? null},
      ${updated.kar ?? null},
      ${updated.m31Degeri ?? null},
      ${updated.veritabaniTeklif ?? null},
      ${updated.license},
      ${existing ? existing.createdAt : now},
      ${now}
    )
    ON CONFLICT (mac_adresi) DO UPDATE SET
      ip_adresi = EXCLUDED.ip_adresi,
      firma_adi = EXCLUDED.firma_adi,
      user_adi = EXCLUDED.user_adi,
      dosya_adi = EXCLUDED.dosya_adi,
      proje_adi = EXCLUDED.proje_adi,
      proje_kisa_adresi = EXCLUDED.proje_kisa_adresi,
      teklif_para_birimi_usd = EXCLUDED.teklif_para_birimi_usd,
      teklif_para_birimi_euro = EXCLUDED.teklif_para_birimi_euro,
      teklif_para_birimi_genel = EXCLUDED.teklif_para_birimi_genel,
      genel_gider = EXCLUDED.genel_gider,
      kar = EXCLUDED.kar,
      m31_degeri = EXCLUDED.m31_degeri,
      veritabani_teklif = EXCLUDED.veritabani_teklif,
      license = EXCLUDED.license,
      updated_at = EXCLUDED.updated_at
    RETURNING
      mac_adresi,
      ip_adresi,
      firma_adi,
      user_adi,
      dosya_adi,
      proje_adi,
      proje_kisa_adresi,
      teklif_para_birimi_usd,
      teklif_para_birimi_euro,
      teklif_para_birimi_genel,
      genel_gider,
      kar,
      m31_degeri,
      veritabani_teklif,
      license,
      created_at,
      updated_at
  `;

  return rowToRecord(rows[0] as LicenseRow);
}

export async function listLicenses(): Promise<LicenseRecord[]> {
  await ensureLicensesTable();
  const sql = getSql();
  const rows = await sql`
    SELECT
      mac_adresi,
      ip_adresi,
      firma_adi,
      user_adi,
      dosya_adi,
      proje_adi,
      proje_kisa_adresi,
      teklif_para_birimi_usd,
      teklif_para_birimi_euro,
      teklif_para_birimi_genel,
      genel_gider,
      kar,
      m31_degeri,
      veritabani_teklif,
      license,
      created_at,
      updated_at
    FROM licenses
    ORDER BY updated_at DESC
  `;

  return (rows as LicenseRow[]).map(rowToRecord);
}

export async function isLicensed(mac: string): Promise<boolean> {
  const record = await getLicenseByMac(mac);
  if (!record) return false;
  const value = record.license.toLowerCase();
  return (
    value === "true" ||
    value === "1" ||
    value === "active" ||
    value === "evet"
  );
}
>>>>>>> 1057c5c83b00dbc2e40b55b442239b0d43d0937c
