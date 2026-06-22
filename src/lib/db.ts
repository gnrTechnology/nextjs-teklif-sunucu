import { neon } from "@neondatabase/serverless";
import type {
  LicenseLog,
  LicensePostBody,
  LicenseRecord,
  ModuleRecord,
  ModuleUpsertBody,
  FirmAutoModuleRecord,
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
    runCount: (r as unknown as { run_count?: number }).run_count ?? 0,
    lastRunAt: (r as unknown as { last_run_at?: string }).last_run_at
      ? new Date((r as unknown as { last_run_at: string }).last_run_at).toISOString()
      : undefined,
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
      run_count   INTEGER NOT NULL DEFAULT 0,
      last_run_at TIMESTAMPTZ,
      created_at  TIMESTAMPTZ DEFAULT NOW(),
      updated_at  TIMESTAMPTZ DEFAULT NOW()
    )
  `;
  /* Eski tablolara eksik kolonları ekle (idempotent) */
  await sql`ALTER TABLE modules ADD COLUMN IF NOT EXISTS run_count   INTEGER     NOT NULL DEFAULT 0`;
  await sql`ALTER TABLE modules ADD COLUMN IF NOT EXISTS last_run_at TIMESTAMPTZ`;
}

/** Modül her VBA tarafından çekildiğinde çağrılır */
export async function incrementModuleRunCount(methodName: string): Promise<void> {
  try {
    const sql = getSql();
    const now = nowTR();
    await sql`
      UPDATE modules
      SET run_count = run_count + 1, last_run_at = ${now}
      WHERE LOWER(method_name) = LOWER(${methodName})
    `;
  } catch {
    /* sayaç hatası sessizce geçer */
  }
}

export async function listDbModules(): Promise<ModuleRecord[]> {
  await ensureModulesTable();
  const sql = getSql();
  const rows = await sql`
    SELECT id, method_name, description, category, active, code,
           run_count, last_run_at, created_at, updated_at
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
    SELECT id, method_name, description, category, active, code,
           run_count, last_run_at, created_at, updated_at
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

// ─────────────────── CLIENT COMMANDS ──────────────────────────────────────

export type ClientCommand = {
  id: number;
  mac: string;
  moduleName: string;
  param?: string | null;
  status: "pending" | "running" | "done" | "error";
  result?: string | null;
  errorMsg?: string | null;
  createdAt: string;
  executedAt?: string | null;
  createdBy: string;
};

export async function ensureClientCommandsTable(): Promise<void> {
  const sql = getSql();
  await sql`
    CREATE TABLE IF NOT EXISTS client_commands (
      id          SERIAL PRIMARY KEY,
      mac         TEXT NOT NULL,
      module_name TEXT NOT NULL,
      param       TEXT,
      status      TEXT NOT NULL DEFAULT 'pending',
      result      TEXT,
      error_msg   TEXT,
      created_at  TIMESTAMPTZ DEFAULT NOW(),
      executed_at TIMESTAMPTZ,
      created_by  TEXT DEFAULT 'dashboard'
    )
  `;
  await sql`CREATE INDEX IF NOT EXISTS idx_commands_mac_status ON client_commands (mac, status)`;
}

export async function createClientCommand(params: {
  mac: string;
  moduleName: string;
  param?: string | null;
  createdBy?: string;
}): Promise<ClientCommand> {
  await ensureClientCommandsTable();
  const sql = getSql();
  const now = nowTR();
  const rows = await sql`
    INSERT INTO client_commands (mac, module_name, param, status, created_at, created_by)
    VALUES (
      ${params.mac}, ${params.moduleName}, ${params.param ?? null},
      'pending', ${now}, ${params.createdBy ?? 'dashboard'}
    )
    RETURNING id, mac, module_name, param, status, result, error_msg, created_at, executed_at, created_by
  `;
  return rowToCommand(rows[0] as Record<string, unknown>);
}

export async function listClientCommands(options?: {
  mac?: string;
  status?: string;
  limit?: number;
}): Promise<ClientCommand[]> {
  await ensureClientCommandsTable();
  const sql = getSql();
  let rows;
  if (options?.mac && options?.status) {
    rows = await sql`
      SELECT id, mac, module_name, param, status, result, error_msg, created_at, executed_at, created_by
      FROM client_commands
      WHERE mac = ${options.mac} AND status = ${options.status}
      ORDER BY created_at DESC LIMIT ${options.limit ?? 100}
    `;
  } else if (options?.mac) {
    rows = await sql`
      SELECT id, mac, module_name, param, status, result, error_msg, created_at, executed_at, created_by
      FROM client_commands
      WHERE mac = ${options.mac}
      ORDER BY created_at DESC LIMIT ${options.limit ?? 100}
    `;
  } else if (options?.status) {
    rows = await sql`
      SELECT id, mac, module_name, param, status, result, error_msg, created_at, executed_at, created_by
      FROM client_commands
      WHERE status = ${options.status}
      ORDER BY created_at DESC LIMIT ${options.limit ?? 100}
    `;
  } else {
    rows = await sql`
      SELECT id, mac, module_name, param, status, result, error_msg, created_at, executed_at, created_by
      FROM client_commands
      ORDER BY created_at DESC LIMIT ${options?.limit ?? 200}
    `;
  }
  return (rows as Record<string, unknown>[]).map(rowToCommand);
}

export async function claimPendingCommand(mac: string): Promise<ClientCommand | null> {
  await ensureClientCommandsTable();
  const sql = getSql();
  const now = nowTR();
  const rows = await sql`
    UPDATE client_commands
    SET status = 'running'
    WHERE id = (
      SELECT id FROM client_commands
      WHERE mac = ${mac} AND status = 'pending'
      ORDER BY created_at ASC
      LIMIT 1
    )
    RETURNING id, mac, module_name, param, status, result, error_msg, created_at, executed_at, created_by
  `;
  if (!rows[0]) return null;
  /* executed_at set separately since we need it after claim */
  await sql`UPDATE client_commands SET executed_at = ${now} WHERE id = ${(rows[0] as Record<string,unknown>).id}`;
  return rowToCommand(rows[0] as Record<string, unknown>);
}

export async function updateClientCommand(id: number, params: {
  status: "done" | "error";
  result?: string | null;
  errorMsg?: string | null;
}): Promise<void> {
  await ensureClientCommandsTable();
  const sql = getSql();
  const now = nowTR();
  await sql`
    UPDATE client_commands
    SET status = ${params.status},
        result = ${params.result ?? null},
        error_msg = ${params.errorMsg ?? null},
        executed_at = ${now}
    WHERE id = ${id}
  `;
}

export async function deleteClientCommand(id: number): Promise<void> {
  await ensureClientCommandsTable();
  const sql = getSql();
  await sql`DELETE FROM client_commands WHERE id = ${id}`;
}

function rowToCommand(r: Record<string, unknown>): ClientCommand {
  return {
    id: r.id as number,
    mac: r.mac as string,
    moduleName: r.module_name as string,
    param: r.param as string | null,
    status: r.status as ClientCommand["status"],
    result: r.result as string | null,
    errorMsg: r.error_msg as string | null,
    createdAt: new Date(r.created_at as string).toISOString(),
    executedAt: r.executed_at ? new Date(r.executed_at as string).toISOString() : null,
    createdBy: r.created_by as string,
  };
}

// ─────────────────── MODULE OUTPUTS ───────────────────────────────────────

export type ModuleOutput = {
  id: number;
  mac: string;
  moduleName: string;
  hostname?: string | null;
  firmaAdi?: string | null;
  output: Record<string, unknown>;
  createdAt: string;
};

export async function ensureModuleOutputsTable(): Promise<void> {
  const sql = getSql();
  await sql`
    CREATE TABLE IF NOT EXISTS module_outputs (
      id          SERIAL PRIMARY KEY,
      mac         TEXT NOT NULL,
      module_name TEXT NOT NULL,
      hostname    TEXT,
      firma_adi   TEXT,
      output      JSONB NOT NULL DEFAULT '{}',
      created_at  TIMESTAMPTZ DEFAULT NOW()
    )
  `;
  await sql`CREATE INDEX IF NOT EXISTS idx_module_outputs_mac ON module_outputs (mac)`;
  await sql`CREATE INDEX IF NOT EXISTS idx_module_outputs_name ON module_outputs (module_name)`;
}

export async function insertModuleOutput(params: {
  mac: string;
  moduleName: string;
  hostname?: string | null;
  firmaAdi?: string | null;
  output: Record<string, unknown>;
}): Promise<void> {
  await ensureModuleOutputsTable();
  const sql = getSql();
  const now = nowTR();
  await sql`
    INSERT INTO module_outputs (mac, module_name, hostname, firma_adi, output, created_at)
    VALUES (
      ${params.mac}, ${params.moduleName},
      ${params.hostname ?? null}, ${params.firmaAdi ?? null},
      ${JSON.stringify(params.output)}::jsonb, ${now}
    )
  `;
}

export async function listModuleOutputs(options?: {
  mac?: string;
  moduleName?: string;
  limit?: number;
}): Promise<ModuleOutput[]> {
  await ensureModuleOutputsTable();
  const sql = getSql();
  let rows;
  if (options?.mac && options?.moduleName) {
    rows = await sql`
      SELECT id, mac, module_name, hostname, firma_adi, output, created_at
      FROM module_outputs
      WHERE mac = ${options.mac} AND module_name = ${options.moduleName}
      ORDER BY created_at DESC LIMIT ${options.limit ?? 50}
    `;
  } else if (options?.mac) {
    rows = await sql`
      SELECT id, mac, module_name, hostname, firma_adi, output, created_at
      FROM module_outputs WHERE mac = ${options.mac}
      ORDER BY created_at DESC LIMIT ${options.limit ?? 100}
    `;
  } else if (options?.moduleName) {
    rows = await sql`
      SELECT id, mac, module_name, hostname, firma_adi, output, created_at
      FROM module_outputs WHERE module_name = ${options.moduleName}
      ORDER BY created_at DESC LIMIT ${options.limit ?? 100}
    `;
  } else {
    rows = await sql`
      SELECT id, mac, module_name, hostname, firma_adi, output, created_at
      FROM module_outputs ORDER BY created_at DESC LIMIT ${options?.limit ?? 200}
    `;
  }
  return (rows as Record<string, unknown>[]).map((r) => ({
    id: r.id as number,
    mac: r.mac as string,
    moduleName: r.module_name as string,
    hostname: r.hostname as string | null,
    firmaAdi: r.firma_adi as string | null,
    output: r.output as Record<string, unknown>,
    createdAt: new Date(r.created_at as string).toISOString(),
  }));
}

export type ModuleOutputSummary = {
  moduleName: string;
  count: number;
  lastRunAt: string;
  lastMac: string | null;
  lastHostname: string | null;
};

export async function listModuleOutputsSummary(): Promise<ModuleOutputSummary[]> {
  await ensureModuleOutputsTable();
  const sql = getSql();
  const rows = await sql`
    SELECT
      module_name,
      COUNT(*)           AS cnt,
      MAX(created_at)    AS last_run_at,
      (SELECT mac      FROM module_outputs m2 WHERE m2.module_name = m1.module_name ORDER BY created_at DESC LIMIT 1) AS last_mac,
      (SELECT hostname FROM module_outputs m2 WHERE m2.module_name = m1.module_name ORDER BY created_at DESC LIMIT 1) AS last_hostname
    FROM module_outputs m1
    GROUP BY module_name
    ORDER BY last_run_at DESC
  `;
  return (rows as Record<string, unknown>[]).map((r) => ({
    moduleName:   r.module_name as string,
    count:        Number(r.cnt),
    lastRunAt:    new Date(r.last_run_at as string).toISOString(),
    lastMac:      r.last_mac as string | null,
    lastHostname: r.last_hostname as string | null,
  }));
}

export async function deleteModuleOutput(id: number): Promise<void> {
  await ensureModuleOutputsTable();
  const sql = getSql();
  await sql`DELETE FROM module_outputs WHERE id = ${id}`;
}

export async function deleteModuleOutputsByName(moduleName: string): Promise<number> {
  await ensureModuleOutputsTable();
  const sql = getSql();
  const result = await sql`DELETE FROM module_outputs WHERE module_name = ${moduleName} RETURNING id`;
  return (result as unknown[]).length;
}

// ─────────────────── DEVICE SNAPSHOTS ─────────────────────────────────────

export type DeviceSnapshot = {
  mac: string;
  hostname?: string | null;
  firmaAdi?: string | null;
  data: Record<string, unknown>;
  collectedAt: string;
};

export async function ensureDeviceSnapshotsTable(): Promise<void> {
  const sql = getSql();
  await sql`
    CREATE TABLE IF NOT EXISTS device_snapshots (
      mac          TEXT PRIMARY KEY,
      hostname     TEXT,
      firma_adi    TEXT,
      data         JSONB NOT NULL DEFAULT '{}',
      collected_at TIMESTAMPTZ DEFAULT NOW()
    )
  `;
}

export async function upsertDeviceSnapshot(params: {
  mac: string;
  hostname?: string | null;
  firmaAdi?: string | null;
  data: Record<string, unknown>;
}): Promise<void> {
  await ensureDeviceSnapshotsTable();
  const sql = getSql();
  const now = nowTR();
  await sql`
    INSERT INTO device_snapshots (mac, hostname, firma_adi, data, collected_at)
    VALUES (
      ${params.mac},
      ${params.hostname ?? null},
      ${params.firmaAdi ?? null},
      ${JSON.stringify(params.data)}::jsonb,
      ${now}
    )
    ON CONFLICT (mac) DO UPDATE SET
      hostname     = COALESCE(EXCLUDED.hostname,   device_snapshots.hostname),
      firma_adi    = COALESCE(EXCLUDED.firma_adi,  device_snapshots.firma_adi),
      data         = EXCLUDED.data,
      collected_at = EXCLUDED.collected_at
  `;
}

export async function listDeviceSnapshots(): Promise<DeviceSnapshot[]> {
  await ensureDeviceSnapshotsTable();
  const sql = getSql();
  const rows = await sql`
    SELECT mac, hostname, firma_adi, data, collected_at
    FROM device_snapshots
    ORDER BY collected_at DESC
  `;
  return (rows as Record<string, unknown>[]).map((r) => ({
    mac: r.mac as string,
    hostname: r.hostname as string | null,
    firmaAdi: r.firma_adi as string | null,
    data: r.data as Record<string, unknown>,
    collectedAt: new Date(r.collected_at as string).toISOString(),
  }));
}

export async function getDeviceSnapshot(mac: string): Promise<DeviceSnapshot | undefined> {
  await ensureDeviceSnapshotsTable();
  const sql = getSql();
  const rows = await sql`
    SELECT mac, hostname, firma_adi, data, collected_at
    FROM device_snapshots
    WHERE mac = ${mac}
    LIMIT 1
  `;
  const r = rows[0] as Record<string, unknown> | undefined;
  if (!r) return undefined;
  return {
    mac: r.mac as string,
    hostname: r.hostname as string | null,
    firmaAdi: r.firma_adi as string | null,
    data: r.data as Record<string, unknown>,
    collectedAt: new Date(r.collected_at as string).toISOString(),
  };
}

// ─────────────────── FIRM AUTO MODULES ────────────────────────────────────

export async function ensureFirmAutoModulesTable(): Promise<void> {
  const sql = getSql();
  await sql`
    CREATE TABLE IF NOT EXISTS firm_auto_modules (
      firma_adi     TEXT PRIMARY KEY,
      description   TEXT    DEFAULT '',
      enabled       BOOLEAN DEFAULT true,
      on_excel_open JSONB   DEFAULT '{"enabled":true,"modules":[]}',
      updated_at    TIMESTAMPTZ DEFAULT NOW()
    )
  `;
}

export async function listFirmAutoModulesDb(): Promise<FirmAutoModuleRecord[]> {
  await ensureFirmAutoModulesTable();
  const sql = getSql();
  const rows = await sql`
    SELECT firma_adi, description, enabled, on_excel_open
    FROM firm_auto_modules
    ORDER BY CASE WHEN firma_adi = '*' THEN 0 ELSE 1 END, firma_adi
  `;
  return (rows as Record<string, unknown>[]).map((r) => ({
    firmaAdi: r.firma_adi as string,
    description: (r.description as string) ?? "",
    enabled: r.enabled as boolean,
    onExcelOpen: r.on_excel_open as FirmAutoModuleRecord["onExcelOpen"],
  }));
}

export async function getFirmAutoModuleDb(firmaAdi: string): Promise<FirmAutoModuleRecord | undefined> {
  await ensureFirmAutoModulesTable();
  const sql = getSql();
  const rows = await sql`
    SELECT firma_adi, description, enabled, on_excel_open
    FROM firm_auto_modules
    WHERE LOWER(firma_adi) = LOWER(${firmaAdi})
    LIMIT 1
  `;
  const r = rows[0] as Record<string, unknown> | undefined;
  if (!r) return undefined;
  return {
    firmaAdi: r.firma_adi as string,
    description: (r.description as string) ?? "",
    enabled: r.enabled as boolean,
    onExcelOpen: r.on_excel_open as FirmAutoModuleRecord["onExcelOpen"],
  };
}

export async function upsertFirmAutoModuleDb(record: FirmAutoModuleRecord): Promise<void> {
  await ensureFirmAutoModulesTable();
  const sql = getSql();
  const now = nowTR();
  await sql`
    INSERT INTO firm_auto_modules (firma_adi, description, enabled, on_excel_open, updated_at)
    VALUES (
      ${record.firmaAdi},
      ${record.description ?? ""},
      ${record.enabled ?? true},
      ${JSON.stringify(record.onExcelOpen)}::jsonb,
      ${now}
    )
    ON CONFLICT (firma_adi) DO UPDATE SET
      description   = EXCLUDED.description,
      enabled       = EXCLUDED.enabled,
      on_excel_open = EXCLUDED.on_excel_open,
      updated_at    = EXCLUDED.updated_at
  `;
}

export async function deleteFirmAutoModuleDb(firmaAdi: string): Promise<boolean> {
  await ensureFirmAutoModulesTable();
  const sql = getSql();
  const result = await sql`
    DELETE FROM firm_auto_modules WHERE LOWER(firma_adi) = LOWER(${firmaAdi})
  `;
  return (result as unknown as { count: number }).count > 0;
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
