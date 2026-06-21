import { neon } from "@neondatabase/serverless";
import type { LicenseLog, LicensePostBody, LicenseRecord } from "./types";

type LicenseRow = {
  mac_adresi: string;
  ip_adresi: string | null;
  firma_adi: string | null;
  user_adi: string | null;
  dosya_adi: string | null;
  license: string;
  created_at: string;
  updated_at: string;
};

function getSql() {
  const databaseUrl = process.env.DATABASE_URL;
  if (!databaseUrl) {
    throw new Error("DATABASE_URL ortam değişkeni tanımlı değil.");
  }
  return neon(databaseUrl);
}

function rowToRecord(row: LicenseRow): LicenseRecord {
  return {
    macAdresi: row.mac_adresi,
    ipAdresi: row.ip_adresi ?? undefined,
    firmaAdi: row.firma_adi ?? undefined,
    userAdi: row.user_adi ?? undefined,
    dosyaAdi: row.dosya_adi ?? undefined,
    license: row.license,
    createdAt: new Date(row.created_at).toISOString(),
    updatedAt: new Date(row.updated_at).toISOString(),
  };
}

const SELECT_COLS = `
  mac_adresi, ip_adresi, firma_adi, user_adi, dosya_adi,
  license, created_at, updated_at
`;

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
  const sql = getSql();
  const normalized = normalizeMac(mac);
  const rows = await sql`
    SELECT ${sql.unsafe(SELECT_COLS)}
    FROM licenses
    WHERE mac_adresi = ${normalized}
    LIMIT 1
  `;
  const row = rows[0] as LicenseRow | undefined;
  return row ? rowToRecord(row) : undefined;
}

/**
 * Tek SQL sorgusu: INSERT ... ON CONFLICT DO UPDATE ... RETURNING
 * xmax = 0 → yeni satır, xmax > 0 → güncelleme
 * COALESCE ile gelen null değerler mevcut değerleri ezmez.
 */
export async function upsertLicense(
  body: LicensePostBody,
): Promise<{ record: LicenseRecord; existed: boolean }> {
  const sql = getSql();
  const normalizedMac = normalizeMac(body.macAdresi);
  const now = new Date().toISOString();

  const rows = await sql`
    INSERT INTO licenses (
      mac_adresi, ip_adresi, firma_adi, user_adi, dosya_adi,
      license, created_at, updated_at
    ) VALUES (
      ${normalizedMac},
      ${body.ipAdresi ?? null},
      ${body.firmaAdi ?? null},
      ${body.userAdi ?? null},
      ${body.dosyaAdi ?? null},
      'false',
      ${now},
      ${now}
    )
    ON CONFLICT (mac_adresi) DO UPDATE SET
      ip_adresi  = COALESCE(EXCLUDED.ip_adresi,  licenses.ip_adresi),
      firma_adi  = COALESCE(EXCLUDED.firma_adi,  licenses.firma_adi),
      user_adi   = COALESCE(EXCLUDED.user_adi,   licenses.user_adi),
      dosya_adi  = COALESCE(EXCLUDED.dosya_adi,  licenses.dosya_adi),
      updated_at = EXCLUDED.updated_at
    RETURNING
      mac_adresi, ip_adresi, firma_adi, user_adi, dosya_adi,
      license, created_at, updated_at,
      (xmax::text::int > 0) AS existed
  `;

  const row = rows[0] as LicenseRow & { existed: boolean };
  return { record: rowToRecord(row), existed: row.existed };
}

export async function listLicenses(): Promise<LicenseRecord[]> {
  const sql = getSql();
  const rows = await sql`
    SELECT ${sql.unsafe(SELECT_COLS)}
    FROM licenses
    ORDER BY updated_at DESC
  `;
  return (rows as LicenseRow[]).map(rowToRecord);
}

export async function isLicensed(mac: string): Promise<boolean> {
  const record = await getLicenseByMac(mac);
  if (!record) return false;
  const value = record.license.toLowerCase();
  return value === "true" || value === "1" || value === "active" || value === "evet";
}

export async function toggleLicense(
  mac: string,
  newValue: "true" | "false",
): Promise<LicenseRecord | null> {
  const sql = getSql();
  const normalized = normalizeMac(mac);
  const now = new Date().toISOString();
  const rows = await sql`
    UPDATE licenses
    SET license = ${newValue}, updated_at = ${now}
    WHERE mac_adresi = ${normalized}
    RETURNING ${sql.unsafe(SELECT_COLS)}
  `;
  const row = rows[0] as LicenseRow | undefined;
  return row ? rowToRecord(row) : null;
}

type InsertLogParams = {
  macAdresi?: string | null;
  firmaAdi?: string | null;
  userAdi?: string | null;
  dosyaAdi?: string | null;
  ipAdresi?: string | null;
  eventType?: string;
  details?: string | null;
};

export async function insertLog(params: InsertLogParams): Promise<void> {
  try {
    const sql = getSql();
    await sql`
      INSERT INTO license_logs
        (mac_adresi, firma_adi, user_adi, dosya_adi, ip_adresi, event_type, details)
      VALUES
        (${params.macAdresi ?? null}, ${params.firmaAdi ?? null}, ${params.userAdi ?? null},
         ${params.dosyaAdi ?? null}, ${params.ipAdresi ?? null},
         ${params.eventType ?? "register"}, ${params.details ?? null})
    `;
  } catch (err) {
    console.error("[insertLog] Log yazılamadı:", err);
  }
}

export async function listLogs(limit = 500): Promise<LicenseLog[]> {
  const sql = getSql();
  const rows = await sql`
    SELECT id, mac_adresi, firma_adi, user_adi, dosya_adi, ip_adresi,
           event_type, details, created_at
    FROM license_logs
    ORDER BY created_at DESC
    LIMIT ${limit}
  `;
  return (rows as Record<string, unknown>[]).map((r) => ({
    id: r.id as number,
    macAdresi: (r.mac_adresi as string) ?? null,
    firmaAdi: (r.firma_adi as string) ?? null,
    userAdi: (r.user_adi as string) ?? null,
    dosyaAdi: (r.dosya_adi as string) ?? null,
    ipAdresi: (r.ip_adresi as string) ?? null,
    eventType: r.event_type as string,
    details: (r.details as string) ?? null,
    createdAt: new Date(r.created_at as string).toISOString(),
  }));
}
