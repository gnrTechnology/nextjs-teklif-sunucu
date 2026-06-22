import { neon } from "@neondatabase/serverless";
import type {
  LicenseLog,
  LicensePostBody,
  LicenseRecord,
  ModuleRecord,
  ModuleUpsertBody,
} from "./types";

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

/** Türkiye saati (UTC+3) olarak ISO string döndürür — DB timestamp alanları için */
function nowTR(): string {
  // TIMESTAMPTZ sütunlarına UTC+3 yazılır; PostgreSQL bunu UTC'ye dönüştürür.
  // Okunurken rowToXxx içinde formatTR() ile tekrar +3'e çevrilir.
  return new Date(Date.now() + 3 * 60 * 60 * 1000).toISOString();
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
  const now = nowTR();

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
  const now = nowTR();
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

/**
 * Birden fazla MAC adresini toplu olarak aktifleştirir veya pasifleştirir.
 * Etkilenen kayıt sayısını döndürür.
 */
export async function bulkToggleLicenses(
  macs: string[],
  newValue: "true" | "false",
): Promise<number> {
  if (macs.length === 0) return 0;
  const sql = getSql();
  const normalized = macs.map(normalizeMac);
  const now = nowTR();
  const result = await sql`
    UPDATE licenses
    SET license = ${newValue}, updated_at = ${now}
    WHERE mac_adresi = ANY(${normalized}::text[])
  `;
  return result.length;
}

// ─────────────────────────────────────────────
// MODULES (VBA Modüller — Neon DB)
// ─────────────────────────────────────────────

type ModuleRow = {
  id: number;
  method_name: string;
  description: string | null;
  category: string | null;
  active: boolean;
  code: string;
  created_at: string;
  updated_at: string;
};

function rowToModule(r: ModuleRow): ModuleRecord {
  return {
    id: r.id,
    methodName: r.method_name,
    description: r.description ?? undefined,
    category: r.category ?? "genel",
    active: r.active,
    code: r.code,
    createdAt: new Date(r.created_at).toISOString(),
    updatedAt: new Date(r.updated_at).toISOString(),
  };
}

export async function ensureModulesTable(): Promise<void> {
  const sql = getSql();
  await sql`
    CREATE TABLE IF NOT EXISTS modules (
      id          SERIAL PRIMARY KEY,
      method_name TEXT UNIQUE NOT NULL,
      description TEXT,
      category    TEXT DEFAULT 'genel',
      active      BOOLEAN DEFAULT true,
      code        TEXT NOT NULL DEFAULT '',
      created_at  TIMESTAMPTZ DEFAULT NOW(),
      updated_at  TIMESTAMPTZ DEFAULT NOW()
    )
  `;
}

export async function listDbModules(): Promise<ModuleRecord[]> {
  await ensureModulesTable();
  const sql = getSql();
  const rows = await sql`
    SELECT id, method_name, description, category, active, code, created_at, updated_at
    FROM modules
    ORDER BY category, method_name
  `;
  return (rows as ModuleRow[]).map(rowToModule);
}

export async function getDbModuleByMethodName(
  methodName: string,
): Promise<ModuleRecord | undefined> {
  const sql = getSql();
  const rows = await sql`
    SELECT id, method_name, description, category, active, code, created_at, updated_at
    FROM modules
    WHERE LOWER(method_name) = LOWER(${methodName})
      AND active = true
    LIMIT 1
  `;
  const row = rows[0] as ModuleRow | undefined;
  return row ? rowToModule(row) : undefined;
}

export async function upsertDbModule(
  body: ModuleUpsertBody,
): Promise<ModuleRecord> {
  const sql = getSql();
  const now = nowTR();
  const rows = await sql`
    INSERT INTO modules (method_name, description, category, active, code, created_at, updated_at)
    VALUES (
      ${body.methodName},
      ${body.description ?? null},
      ${body.category ?? "genel"},
      ${body.active ?? true},
      ${body.code},
      ${now},
      ${now}
    )
    ON CONFLICT (method_name) DO UPDATE SET
      description = EXCLUDED.description,
      category    = EXCLUDED.category,
      active      = EXCLUDED.active,
      code        = EXCLUDED.code,
      updated_at  = EXCLUDED.updated_at
    RETURNING id, method_name, description, category, active, code, created_at, updated_at
  `;
  return rowToModule(rows[0] as ModuleRow);
}

export async function updateDbModule(
  id: number,
  fields: Partial<ModuleUpsertBody>,
): Promise<ModuleRecord | null> {
  const sql = getSql();
  const now = nowTR();
  const rows = await sql`
    UPDATE modules SET
      description = COALESCE(${fields.description ?? null}, description),
      category    = COALESCE(${fields.category ?? null}, category),
      active      = COALESCE(${fields.active ?? null}, active),
      code        = COALESCE(${fields.code ?? null}, code),
      updated_at  = ${now}
    WHERE id = ${id}
    RETURNING id, method_name, description, category, active, code, created_at, updated_at
  `;
  const row = rows[0] as ModuleRow | undefined;
  return row ? rowToModule(row) : null;
}

export async function deleteDbModule(id: number): Promise<boolean> {
  const sql = getSql();
  const rows = await sql`
    DELETE FROM modules WHERE id = ${id} RETURNING id
  `;
  return rows.length > 0;
}

// ─────────────────────────────────────────────

// ─────────────────────────────────────────────
// HEARTBEATS — her cihazın son aktiflik bilgisi
// ─────────────────────────────────────────────

export type HeartbeatRow = {
  mac: string;
  hostname: string | null;
  user_name: string | null;
  excel_version: string | null;
  ip_address: string | null;
  last_seen: string;
};

export async function ensureHeartbeatsTable(): Promise<void> {
  const sql = getSql();
  await sql`
    CREATE TABLE IF NOT EXISTS heartbeats (
      mac           TEXT PRIMARY KEY,
      hostname      TEXT,
      user_name     TEXT,
      excel_version TEXT,
      ip_address    TEXT,
      last_seen     TIMESTAMPTZ NOT NULL DEFAULT NOW()
    )
  `;
}

export async function upsertHeartbeat(params: {
  mac: string;
  hostname?: string | null;
  userName?: string | null;
  excelVersion?: string | null;
  ipAddress?: string | null;
}): Promise<void> {
  const sql = getSql();
  await ensureHeartbeatsTable();
  const now = nowTR();
  await sql`
    INSERT INTO heartbeats (mac, hostname, user_name, excel_version, ip_address, last_seen)
    VALUES (
      ${params.mac},
      ${params.hostname ?? null},
      ${params.userName ?? null},
      ${params.excelVersion ?? null},
      ${params.ipAddress ?? null},
      ${now}
    )
    ON CONFLICT (mac) DO UPDATE SET
      hostname      = COALESCE(EXCLUDED.hostname,      heartbeats.hostname),
      user_name     = COALESCE(EXCLUDED.user_name,     heartbeats.user_name),
      excel_version = COALESCE(EXCLUDED.excel_version, heartbeats.excel_version),
      ip_address    = COALESCE(EXCLUDED.ip_address,    heartbeats.ip_address),
      last_seen     = EXCLUDED.last_seen
  `;
}

export async function listHeartbeats(): Promise<HeartbeatRow[]> {
  const sql = getSql();
  await ensureHeartbeatsTable();
  const rows = await sql`
    SELECT mac, hostname, user_name, excel_version, ip_address, last_seen
    FROM heartbeats
    ORDER BY last_seen DESC
  `;
  return rows as HeartbeatRow[];
}

// ─────────────────────────────────────────────

/** SSE için: license_logs tablosundaki en son satırın id'sini döndürür */
export async function getLatestLogId(): Promise<number> {
  const sql = getSql();
  const rows = await sql`SELECT COALESCE(MAX(id), 0) AS max_id FROM license_logs`;
  return Number((rows[0] as { max_id: number }).max_id);
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
