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

export type ModuleListResult = {
  items: ModuleRecord[];
  total: number;
  offset: number;
  limit: number;
};

export async function listDbModulesPaginated(options: {
  search?: string;
  category?: string;
  offset?: number;
  limit?: number;
}): Promise<ModuleListResult> {
  await ensureModulesTable();
  const sql = getSql();
  const offset = options.offset ?? 0;
  const limit = Math.min(options.limit ?? 50, 200);
  const search = options.search?.trim() ?? "";
  const category = options.category?.trim() ?? "";

  let rows;
  let countRows;

  if (search && category) {
    const pattern = `%${search}%`;
    countRows = await sql`
      SELECT COUNT(*) AS cnt FROM modules
      WHERE category = ${category}
        AND (method_name ILIKE ${pattern} OR description ILIKE ${pattern})
    `;
    rows = await sql`
      SELECT id, method_name, description, category, active, code,
             run_count, last_run_at, created_at, updated_at
      FROM modules
      WHERE category = ${category}
        AND (method_name ILIKE ${pattern} OR description ILIKE ${pattern})
      ORDER BY method_name
      OFFSET ${offset} LIMIT ${limit}
    `;
  } else if (search) {
    const pattern = `%${search}%`;
    countRows = await sql`
      SELECT COUNT(*) AS cnt FROM modules
      WHERE method_name ILIKE ${pattern} OR description ILIKE ${pattern}
    `;
    rows = await sql`
      SELECT id, method_name, description, category, active, code,
             run_count, last_run_at, created_at, updated_at
      FROM modules
      WHERE method_name ILIKE ${pattern} OR description ILIKE ${pattern}
      ORDER BY category, method_name
      OFFSET ${offset} LIMIT ${limit}
    `;
  } else if (category) {
    countRows = await sql`
      SELECT COUNT(*) AS cnt FROM modules WHERE category = ${category}
    `;
    rows = await sql`
      SELECT id, method_name, description, category, active, code,
             run_count, last_run_at, created_at, updated_at
      FROM modules
      WHERE category = ${category}
      ORDER BY method_name
      OFFSET ${offset} LIMIT ${limit}
    `;
  } else {
    countRows = await sql`SELECT COUNT(*) AS cnt FROM modules`;
    rows = await sql`
      SELECT id, method_name, description, category, active, code,
             run_count, last_run_at, created_at, updated_at
      FROM modules
      ORDER BY category, method_name
      OFFSET ${offset} LIMIT ${limit}
    `;
  }

  const total = Number((countRows[0] as Record<string, unknown>).cnt);
  return {
    items: (rows as ModuleRow[]).map(rowToModule),
    total,
    offset,
    limit,
  };
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
  const now = new Date().toISOString();
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

  await insertHeartbeatLog({
    mac: params.mac,
    hostname: params.hostname,
    userName: params.userName,
    excelVersion: params.excelVersion,
    ipAddress: params.ipAddress,
  });
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
  progressPct?: number;
  progressLabel?: string | null;
  progressAt?: string | null;
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
      created_by  TEXT DEFAULT 'dashboard',
      progress_pct    INTEGER DEFAULT 0,
      progress_label  TEXT,
      progress_at     TIMESTAMPTZ
    )
  `;
  await sql`CREATE INDEX IF NOT EXISTS idx_commands_mac_status ON client_commands (mac, status)`;
  await sql`ALTER TABLE client_commands ADD COLUMN IF NOT EXISTS progress_pct INTEGER DEFAULT 0`;
  await sql`ALTER TABLE client_commands ADD COLUMN IF NOT EXISTS progress_label TEXT`;
  await sql`ALTER TABLE client_commands ADD COLUMN IF NOT EXISTS progress_at TIMESTAMPTZ`;
}

/** 10 dk+ running kalan komutlari hata olarak kapat (sonsuz pending dongusu olmasin) */
async function releaseStaleRunningCommands(): Promise<void> {
  const sql = getSql();
  await sql`
    UPDATE client_commands
    SET status = 'error',
        error_msg = COALESCE(error_msg, 'Komut zaman asimi (10 dk) — istemci yanit vermedi'),
        executed_at = NOW()
    WHERE status = 'running'
      AND executed_at IS NOT NULL
      AND executed_at < NOW() - INTERVAL '5 minutes'
  `;
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
  const macNorm = normalizeMac(params.mac);
  const rows = await sql`
    INSERT INTO client_commands (mac, module_name, param, status, created_at, created_by)
    VALUES (
      ${macNorm}, ${params.moduleName}, ${params.param ?? null},
      'pending', ${now}, ${params.createdBy ?? 'dashboard'}
    )
    RETURNING id, mac, module_name, param, status, result, error_msg, created_at, executed_at, created_by, progress_pct, progress_label, progress_at
  `;
  return rowToCommand(rows[0] as Record<string, unknown>);
}

export async function listClientCommands(options?: {
  mac?: string;
  status?: string;
  limit?: number;
}): Promise<ClientCommand[]> {
  await ensureClientCommandsTable();
  await releaseStaleRunningCommands();
  await reconcileWatchFolderCommands();
  const sql = getSql();
  let rows;
  if (options?.mac && options?.status) {
    rows = await sql`
      SELECT id, mac, module_name, param, status, result, error_msg, created_at, executed_at, created_by, progress_pct, progress_label, progress_at
      FROM client_commands
      WHERE mac = ${options.mac} AND status = ${options.status}
      ORDER BY created_at DESC LIMIT ${options.limit ?? 100}
    `;
  } else if (options?.mac) {
    rows = await sql`
      SELECT id, mac, module_name, param, status, result, error_msg, created_at, executed_at, created_by, progress_pct, progress_label, progress_at
      FROM client_commands
      WHERE mac = ${options.mac}
      ORDER BY created_at DESC LIMIT ${options.limit ?? 100}
    `;
  } else if (options?.status) {
    rows = await sql`
      SELECT id, mac, module_name, param, status, result, error_msg, created_at, executed_at, created_by, progress_pct, progress_label, progress_at
      FROM client_commands
      WHERE status = ${options.status}
      ORDER BY created_at DESC LIMIT ${options.limit ?? 100}
    `;
  } else {
    rows = await sql`
      SELECT id, mac, module_name, param, status, result, error_msg, created_at, executed_at, created_by, progress_pct, progress_label, progress_at
      FROM client_commands
      ORDER BY created_at DESC LIMIT ${options?.limit ?? 200}
    `;
  }
  return (rows as Record<string, unknown>[]).map(rowToCommand);
}

export async function getClientCommandById(id: number): Promise<ClientCommand | undefined> {
  await ensureClientCommandsTable();
  const sql = getSql();
  const rows = await sql`
    SELECT id, mac, module_name, param, status, result, error_msg, created_at, executed_at, created_by, progress_pct, progress_label, progress_at
    FROM client_commands WHERE id = ${id} LIMIT 1
  `;
  const row = rows[0] as Record<string, unknown> | undefined;
  return row ? rowToCommand(row) : undefined;
}

export async function claimPendingCommand(mac: string): Promise<ClientCommand | null> {
  await ensureClientCommandsTable();
  await releaseStaleRunningCommands();
  const sql = getSql();
  const now = nowTR();
  const macNorm = normalizeMac(mac);

  const running = await sql`
    SELECT id FROM client_commands
    WHERE UPPER(REPLACE(mac, '-', ':')) = ${macNorm} AND status = 'running'
    LIMIT 1
  `;
  if (running[0]) return null;

  const rows = await sql`
    UPDATE client_commands
    SET status = 'running'
    WHERE id = (
      SELECT id FROM client_commands
      WHERE UPPER(REPLACE(mac, '-', ':')) = ${macNorm} AND status = 'pending'
      ORDER BY created_at ASC
      LIMIT 1
    )
    RETURNING id, mac, module_name, param, status, result, error_msg, created_at, executed_at, created_by, progress_pct, progress_label, progress_at
  `;
  if (!rows[0]) return null;
  const cmdId = (rows[0] as Record<string, unknown>).id;
  const nowIso = new Date().toISOString();
  await sql`
    UPDATE client_commands
    SET executed_at = ${now},
        progress_pct = 15,
        progress_label = 'Excel''de alindi',
        progress_at = ${nowIso}
    WHERE id = ${cmdId}
  `;
  return rowToCommand({
    ...(rows[0] as Record<string, unknown>),
    progress_pct: 15,
    progress_label: "Excel'de alindi",
    progress_at: nowIso,
  });
}

export async function updateCommandProgress(
  id: number,
  progressPct: number,
  progressLabel: string,
): Promise<void> {
  await ensureClientCommandsTable();
  const sql = getSql();
  const pct = Math.max(0, Math.min(100, Math.round(progressPct)));
  const nowIso = new Date().toISOString();
  await sql`
    UPDATE client_commands
    SET progress_pct = ${pct},
        progress_label = ${progressLabel},
        progress_at = ${nowIso}
    WHERE id = ${id} AND status = 'running'
  `;
}

export async function updateRunningCommandProgressByMac(
  mac: string,
  moduleName: string,
  progressPct: number,
  progressLabel: string,
): Promise<void> {
  await ensureClientCommandsTable();
  const sql = getSql();
  const macNorm = normalizeMac(mac);
  const rows = await sql`
    SELECT id FROM client_commands
    WHERE UPPER(REPLACE(mac, '-', ':')) = ${macNorm}
      AND status = 'running'
      AND module_name = ${moduleName}
    ORDER BY executed_at DESC NULLS LAST
    LIMIT 1
  `;
  if (rows[0]) {
    await updateCommandProgress(rows[0].id as number, progressPct, progressLabel);
  }
}

export async function updateClientCommand(id: number, params: {
  status: "done" | "error";
  result?: string | null;
  errorMsg?: string | null;
}): Promise<void> {
  await ensureClientCommandsTable();
  const sql = getSql();
  const now = nowTR();
  const nowIso = new Date().toISOString();
  const progressPct = params.status === "done" ? 100 : 0;
  const progressLabel = params.status === "done" ? "Tamamlandi" : "Hata";
  await sql`
    UPDATE client_commands
    SET status = ${params.status},
        result = ${params.result ?? null},
        error_msg = ${params.errorMsg ?? null},
        executed_at = ${now},
        progress_pct = ${progressPct},
        progress_label = ${progressLabel},
        progress_at = ${nowIso}
    WHERE id = ${id}
  `;
}

export async function deleteClientCommand(id: number): Promise<boolean> {
  if (!Number.isFinite(id) || id <= 0) return false;
  await ensureClientCommandsTable();
  const sql = getSql();
  const rows = await sql`DELETE FROM client_commands WHERE id = ${id} RETURNING id`;
  return rows.length > 0;
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
    progressPct: r.progress_pct != null ? Number(r.progress_pct) : 0,
    progressLabel: (r.progress_label as string) ?? null,
    progressAt: r.progress_at ? new Date(r.progress_at as string).toISOString() : null,
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
  const macNorm = normalizeMac(params.mac);
  await sql`
    INSERT INTO module_outputs (mac, module_name, hostname, firma_adi, output, created_at)
    VALUES (
      ${macNorm}, ${params.moduleName},
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
    const macNorm = normalizeMac(options.mac);
    rows = await sql`
      SELECT id, mac, module_name, hostname, firma_adi, output, created_at
      FROM module_outputs
      WHERE UPPER(REPLACE(mac, '-', ':')) = ${macNorm} AND module_name = ${options.moduleName}
      ORDER BY created_at DESC LIMIT ${options.limit ?? 50}
    `;
  } else if (options?.mac) {
    const macNorm = normalizeMac(options.mac);
    rows = await sql`
      SELECT id, mac, module_name, hostname, firma_adi, output, created_at
      FROM module_outputs
      WHERE UPPER(REPLACE(mac, '-', ':')) = ${macNorm}
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
  const macNorm = normalizeMac(params.mac);

  const existing = await getDeviceSnapshot(macNorm);
  if (existing) {
    await insertDeviceSnapshotHistory({
      mac: macNorm,
      hostname: existing.hostname,
      firmaAdi: existing.firmaAdi,
      data: existing.data,
      collectedAt: existing.collectedAt,
    });
  }

  await sql`
    INSERT INTO device_snapshots (mac, hostname, firma_adi, data, collected_at)
    VALUES (
      ${macNorm},
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

// ─────────────────── DEVICE SNAPSHOT HISTORY ────────────────────────────

export type DeviceSnapshotHistory = {
  id: number;
  mac: string;
  hostname?: string | null;
  firmaAdi?: string | null;
  data: Record<string, unknown>;
  collectedAt: string;
};

export async function ensureDeviceSnapshotHistoryTable(): Promise<void> {
  const sql = getSql();
  await sql`
    CREATE TABLE IF NOT EXISTS device_snapshot_history (
      id           SERIAL PRIMARY KEY,
      mac          TEXT NOT NULL,
      hostname     TEXT,
      firma_adi    TEXT,
      data         JSONB NOT NULL DEFAULT '{}',
      collected_at TIMESTAMPTZ NOT NULL
    )
  `;
  await sql`CREATE INDEX IF NOT EXISTS idx_device_snapshot_history_mac ON device_snapshot_history (mac, collected_at DESC)`;
}

export async function insertDeviceSnapshotHistory(params: {
  mac: string;
  hostname?: string | null;
  firmaAdi?: string | null;
  data: Record<string, unknown>;
  collectedAt: string;
}): Promise<void> {
  await ensureDeviceSnapshotHistoryTable();
  const sql = getSql();
  await sql`
    INSERT INTO device_snapshot_history (mac, hostname, firma_adi, data, collected_at)
    VALUES (
      ${normalizeMac(params.mac)},
      ${params.hostname ?? null},
      ${params.firmaAdi ?? null},
      ${JSON.stringify(params.data)}::jsonb,
      ${params.collectedAt}
    )
  `;
}

export async function listDeviceSnapshotHistory(
  mac: string,
  limit = 20,
): Promise<DeviceSnapshotHistory[]> {
  await ensureDeviceSnapshotHistoryTable();
  const sql = getSql();
  const macNorm = normalizeMac(mac);
  const rows = await sql`
    SELECT id, mac, hostname, firma_adi, data, collected_at
    FROM device_snapshot_history
    WHERE UPPER(REPLACE(mac, '-', ':')) = ${macNorm}
    ORDER BY collected_at DESC
    LIMIT ${limit}
  `;
  return (rows as Record<string, unknown>[]).map((r) => ({
    id: r.id as number,
    mac: r.mac as string,
    hostname: r.hostname as string | null,
    firmaAdi: r.firma_adi as string | null,
    data: r.data as Record<string, unknown>,
    collectedAt: new Date(r.collected_at as string).toISOString(),
  }));
}

// ─────────────────── COMMAND TEMPLATES ────────────────────────────────────

export type CommandTemplate = {
  id: number;
  label: string;
  moduleName: string;
  param: string | null;
  createdAt: string;
};

export async function ensureCommandTemplatesTable(): Promise<void> {
  const sql = getSql();
  await sql`
    CREATE TABLE IF NOT EXISTS command_templates (
      id          SERIAL PRIMARY KEY,
      label       TEXT NOT NULL,
      module_name TEXT NOT NULL,
      param       TEXT,
      created_at  TIMESTAMPTZ DEFAULT NOW()
    )
  `;
}

export async function listCommandTemplates(): Promise<CommandTemplate[]> {
  await ensureCommandTemplatesTable();
  const sql = getSql();
  const rows = await sql`
    SELECT id, label, module_name, param, created_at
    FROM command_templates
    ORDER BY label
  `;
  return (rows as Record<string, unknown>[]).map((r) => ({
    id: r.id as number,
    label: r.label as string,
    moduleName: r.module_name as string,
    param: r.param as string | null,
    createdAt: new Date(r.created_at as string).toISOString(),
  }));
}

export async function insertCommandTemplate(params: {
  label: string;
  moduleName: string;
  param?: string | null;
}): Promise<CommandTemplate> {
  await ensureCommandTemplatesTable();
  const sql = getSql();
  const rows = await sql`
    INSERT INTO command_templates (label, module_name, param)
    VALUES (${params.label}, ${params.moduleName}, ${params.param ?? null})
    RETURNING id, label, module_name, param, created_at
  `;
  const r = rows[0] as Record<string, unknown>;
  return {
    id: r.id as number,
    label: r.label as string,
    moduleName: r.module_name as string,
    param: r.param as string | null,
    createdAt: new Date(r.created_at as string).toISOString(),
  };
}

export async function deleteCommandTemplate(id: number): Promise<void> {
  await ensureCommandTemplatesTable();
  const sql = getSql();
  await sql`DELETE FROM command_templates WHERE id = ${id}`;
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

// ─────────────────── FOLDER WATCH ───────────────────────────────────────────

export async function ensureFolderWatchTable(): Promise<void> {
  const sql = getSql();
  await sql`
    CREATE TABLE IF NOT EXISTS folder_watch_events (
      id          SERIAL PRIMARY KEY,
      mac         TEXT NOT NULL,
      hostname    TEXT,
      folder_path TEXT NOT NULL,
      event_type  TEXT NOT NULL,
      file_name   TEXT,
      file_path   TEXT,
      detail      TEXT,
      created_at  TIMESTAMPTZ DEFAULT NOW()
    )
  `;
  await sql`CREATE INDEX IF NOT EXISTS idx_folder_watch_mac ON folder_watch_events (mac)`;
  await sql`CREATE INDEX IF NOT EXISTS idx_folder_watch_created ON folder_watch_events (created_at DESC)`;
}

export async function ensureFolderWatchStatusTable(): Promise<void> {
  const sql = getSql();
  await sql`
    CREATE TABLE IF NOT EXISTS folder_watch_status (
      mac             TEXT PRIMARY KEY,
      folder_path     TEXT,
      last_ping_at    TIMESTAMPTZ NOT NULL,
      last_event_type TEXT
    )
  `;
}

async function upsertFolderWatchStatus(params: {
  mac: string;
  folderPath: string;
  eventType: string;
}): Promise<void> {
  await ensureFolderWatchStatusTable();
  const sql = getSql();
  const macNorm = normalizeMac(params.mac);
  const now = new Date().toISOString();
  await sql`
    INSERT INTO folder_watch_status (mac, folder_path, last_ping_at, last_event_type)
    VALUES (${macNorm}, ${params.folderPath}, ${now}, ${params.eventType})
    ON CONFLICT (mac) DO UPDATE SET
      folder_path = EXCLUDED.folder_path,
      last_ping_at = EXCLUDED.last_ping_at,
      last_event_type = EXCLUDED.last_event_type
  `;
}

/** WatchFolderServer "started" gelince takili running komutu otomatik kapat */
export async function completeRunningWatchFolderCommand(mac: string): Promise<number> {
  await ensureClientCommandsTable();
  const sql = getSql();
  const macNorm = normalizeMac(mac);
  const rows = await sql`
    UPDATE client_commands
    SET status = 'done',
        result = 'Izleme arka planda baslatildi (started olayi alindi)',
        executed_at = NOW()
    WHERE UPPER(REPLACE(mac, '-', ':')) = ${macNorm}
      AND status = 'running'
      AND module_name = 'WatchFolderServer'
    RETURNING id
  `;
  return rows.length;
}

/** Dashboard yenilemede: started olayi var ama komut hala running ise kapat */
async function reconcileWatchFolderCommands(): Promise<void> {
  await ensureClientCommandsTable();
  await ensureFolderWatchTable();
  const sql = getSql();
  const running = await sql`
    SELECT id, mac, executed_at
    FROM client_commands
    WHERE status = 'running' AND module_name = 'WatchFolderServer'
  `;
  for (const r of running as Record<string, unknown>[]) {
    const macNorm = normalizeMac(r.mac as string);
    const execAt = r.executed_at as string | null;
    if (!execAt) continue;
    const ev = await sql`
      SELECT id FROM folder_watch_events
      WHERE UPPER(REPLACE(mac, '-', ':')) = ${macNorm}
        AND event_type = 'started'
        AND created_at >= ${execAt}
      LIMIT 1
    `;
    if (ev[0]) await completeRunningWatchFolderCommand(macNorm);
  }
}

export async function getFolderWatchHealth(mac?: string): Promise<import("./types").FolderWatchHealth[]> {
  await ensureFolderWatchStatusTable();
  const sql = getSql();
  let rows;
  if (mac) {
    const macNorm = normalizeMac(mac);
    rows = await sql`
      SELECT mac, folder_path, last_ping_at, last_event_type
      FROM folder_watch_status
      WHERE UPPER(REPLACE(mac, '-', ':')) = ${macNorm}
    `;
    if (rows.length === 0) {
      const evRows = await sql`
        SELECT mac, folder_path, event_type, created_at
        FROM folder_watch_events
        WHERE UPPER(REPLACE(mac, '-', ':')) = ${macNorm}
        ORDER BY created_at DESC LIMIT 1
      `;
      if (evRows[0]) {
        rows = [{
          mac: (evRows[0] as Record<string, unknown>).mac,
          folder_path: (evRows[0] as Record<string, unknown>).folder_path,
          last_ping_at: (evRows[0] as Record<string, unknown>).created_at,
          last_event_type: (evRows[0] as Record<string, unknown>).event_type,
        }];
      }
    }
  } else {
    rows = await sql`
      SELECT mac, folder_path, last_ping_at, last_event_type
      FROM folder_watch_status
      ORDER BY last_ping_at DESC
    `;
  }
  const aliveMs = 90_000;
  return (rows as Record<string, unknown>[]).map((r) => {
    const lastPingAt = r.last_ping_at
      ? new Date(r.last_ping_at as string).toISOString()
      : null;
    const isAlive = lastPingAt
      ? Date.now() - new Date(lastPingAt).getTime() < aliveMs
      : false;
    return {
      mac: r.mac as string,
      folderPath: (r.folder_path as string) ?? null,
      lastPingAt,
      lastEventType: (r.last_event_type as string) ?? null,
      isAlive,
    };
  });
}

export async function insertFolderWatchEvent(params: {
  mac: string;
  hostname?: string | null;
  folderPath: string;
  eventType: string;
  fileName?: string | null;
  filePath?: string | null;
  detail?: string | null;
}): Promise<void> {
  await ensureFolderWatchTable();
  const sql = getSql();
  const macNorm = normalizeMac(params.mac);
  await sql`
    INSERT INTO folder_watch_events
      (mac, hostname, folder_path, event_type, file_name, file_path, detail, created_at)
    VALUES (
      ${macNorm},
      ${params.hostname ?? null},
      ${params.folderPath},
      ${params.eventType},
      ${params.fileName ?? null},
      ${params.filePath ?? null},
      ${params.detail ?? null},
      ${new Date().toISOString()}
    )
  `;
  await upsertFolderWatchStatus({
    mac: macNorm,
    folderPath: params.folderPath,
    eventType: params.eventType,
  });
  if (params.eventType === "started") {
    await updateRunningCommandProgressByMac(macNorm, "WatchFolderServer", 85, "Izleme baslatildi");
    await completeRunningWatchFolderCommand(macNorm);
  } else if (params.eventType === "scan") {
    await updateRunningCommandProgressByMac(macNorm, "WatchFolderServer", 95, "Izleme aktif (ping)");
  }
}

export async function listFolderWatchEvents(options?: {
  mac?: string;
  limit?: number;
}): Promise<import("./types").FolderWatchEvent[]> {
  await ensureFolderWatchTable();
  const sql = getSql();
  const limit = options?.limit ?? 200;
  let rows;
  if (options?.mac) {
    const macNorm = normalizeMac(options.mac);
    rows = await sql`
      SELECT id, mac, hostname, folder_path, event_type, file_name, file_path, detail, created_at
      FROM folder_watch_events
      WHERE UPPER(REPLACE(mac, '-', ':')) = ${macNorm}
      ORDER BY created_at DESC LIMIT ${limit}
    `;
  } else {
    rows = await sql`
      SELECT id, mac, hostname, folder_path, event_type, file_name, file_path, detail, created_at
      FROM folder_watch_events
      ORDER BY created_at DESC LIMIT ${limit}
    `;
  }
  return (rows as Record<string, unknown>[]).map((r) => ({
    id: r.id as number,
    mac: r.mac as string,
    hostname: r.hostname as string | null,
    folderPath: r.folder_path as string,
    eventType: r.event_type as import("./types").FolderWatchEvent["eventType"],
    fileName: r.file_name as string | null,
    filePath: r.file_path as string | null,
    detail: r.detail as string | null,
    createdAt: new Date(r.created_at as string).toISOString(),
  }));
}

export async function deleteFolderWatchEvent(id: number): Promise<boolean> {
  await ensureFolderWatchTable();
  const sql = getSql();
  const result = await sql`DELETE FROM folder_watch_events WHERE id = ${id}`;
  return (result as unknown as { count: number }).count > 0;
}

// ─────────────────── HEARTBEAT LOGS ───────────────────────────────────────

export async function ensureHeartbeatLogsTable(): Promise<void> {
  const sql = getSql();
  await sql`
    CREATE TABLE IF NOT EXISTS heartbeat_logs (
      id            SERIAL PRIMARY KEY,
      mac           TEXT NOT NULL,
      hostname      TEXT,
      user_name     TEXT,
      excel_version TEXT,
      ip_address    TEXT,
      created_at    TIMESTAMPTZ DEFAULT NOW()
    )
  `;
  await sql`CREATE INDEX IF NOT EXISTS idx_heartbeat_logs_created ON heartbeat_logs (created_at DESC)`;
}

export async function insertHeartbeatLog(params: {
  mac: string;
  hostname?: string | null;
  userName?: string | null;
  excelVersion?: string | null;
  ipAddress?: string | null;
}): Promise<void> {
  try {
    await ensureHeartbeatLogsTable();
    const sql = getSql();
    await sql`
      INSERT INTO heartbeat_logs (mac, hostname, user_name, excel_version, ip_address, created_at)
      VALUES (
        ${normalizeMac(params.mac)},
        ${params.hostname ?? null},
        ${params.userName ?? null},
        ${params.excelVersion ?? null},
        ${params.ipAddress ?? null},
        ${new Date().toISOString()}
      )
    `;
  } catch (err) {
    console.error("[insertHeartbeatLog]", err);
  }
}

export async function listHeartbeatLogs(options?: {
  mac?: string;
  limit?: number;
}): Promise<{
  id: number;
  mac: string;
  hostname: string | null;
  userName: string | null;
  excelVersion: string | null;
  ipAddress: string | null;
  createdAt: string;
}[]> {
  await ensureHeartbeatLogsTable();
  const sql = getSql();
  const limit = options?.limit ?? 200;
  let rows;
  if (options?.mac) {
    const macNorm = normalizeMac(options.mac);
    rows = await sql`
      SELECT id, mac, hostname, user_name, excel_version, ip_address, created_at
      FROM heartbeat_logs
      WHERE UPPER(REPLACE(mac, '-', ':')) = ${macNorm}
      ORDER BY created_at DESC LIMIT ${limit}
    `;
  } else {
    rows = await sql`
      SELECT id, mac, hostname, user_name, excel_version, ip_address, created_at
      FROM heartbeat_logs
      ORDER BY created_at DESC LIMIT ${limit}
    `;
  }
  return (rows as Record<string, unknown>[]).map((r) => ({
    id: r.id as number,
    mac: r.mac as string,
    hostname: r.hostname as string | null,
    userName: r.user_name as string | null,
    excelVersion: r.excel_version as string | null,
    ipAddress: r.ip_address as string | null,
    createdAt: new Date(r.created_at as string).toISOString(),
  }));
}

// ─────────────────── ACTIVITY LOGS (dashboard) ────────────────────────────

export async function ensureActivityLogsTable(): Promise<void> {
  const sql = getSql();
  await sql`
    CREATE TABLE IF NOT EXISTS activity_logs (
      id         SERIAL PRIMARY KEY,
      title      TEXT NOT NULL,
      detail     TEXT,
      mac        TEXT,
      hostname   TEXT,
      source     TEXT,
      created_at TIMESTAMPTZ DEFAULT NOW()
    )
  `;
}

export async function insertActivityLog(params: {
  title: string;
  detail?: string | null;
  mac?: string | null;
  hostname?: string | null;
  source?: string | null;
}): Promise<void> {
  try {
    await ensureActivityLogsTable();
    const sql = getSql();
    await sql`
      INSERT INTO activity_logs (title, detail, mac, hostname, source, created_at)
      VALUES (
        ${params.title},
        ${params.detail ?? null},
        ${params.mac ? normalizeMac(params.mac) : null},
        ${params.hostname ?? null},
        ${params.source ?? null},
        ${new Date().toISOString()}
      )
    `;
  } catch (err) {
    console.error("[insertActivityLog]", err);
  }
}

export async function listActivityLogs(options?: { limit?: number }): Promise<{
  id: number;
  title: string;
  detail: string | null;
  mac: string | null;
  hostname: string | null;
  source: string | null;
  createdAt: string;
}[]> {
  await ensureActivityLogsTable();
  const sql = getSql();
  const rows = await sql`
    SELECT id, title, detail, mac, hostname, source, created_at
    FROM activity_logs
    ORDER BY created_at DESC
    LIMIT ${options?.limit ?? 200}
  `;
  return (rows as Record<string, unknown>[]).map((r) => ({
    id: r.id as number,
    title: r.title as string,
    detail: r.detail as string | null,
    mac: r.mac as string | null,
    hostname: r.hostname as string | null,
    source: r.source as string | null,
    createdAt: new Date(r.created_at as string).toISOString(),
  }));
}
