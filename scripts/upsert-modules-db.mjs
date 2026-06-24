/**
 * modules tablosuna doğrudan Neon upsert — /api/modules/seed kullanmadan.
 *
 * Kullanım:
 *   node scripts/upsert-modules-db.mjs InstallCommandQueue WatchFolderServer
 *   node scripts/upsert-modules-db.mjs --all-json
 *   node scripts/upsert-modules-db.mjs --all-bas
 */
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import { neon } from "@neondatabase/serverless";

const root = path.join(path.dirname(fileURLToPath(import.meta.url)), "..");
const sourceDir = path.join(root, "data", "modules-source");
const jsonPath = path.join(root, "data", "modules.json");
const manifestPath = path.join(root, "data", "modules.manifest.json");

const SKIP_BAS = new Set([
  "zInternet-additions.bas",
  "TeklifBootstrap.bas",
  "FolderWatchPollHelpers.bas",
]);

function loadEnvLocal() {
  const envPath = path.join(root, ".env.local");
  if (!fs.existsSync(envPath)) {
    throw new Error(".env.local bulunamadı (DATABASE_URL gerekli).");
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

function readJson(filePath) {
  if (!fs.existsSync(filePath)) return [];
  try {
    const data = JSON.parse(fs.readFileSync(filePath, "utf-8"));
    return Array.isArray(data) ? data : [];
  } catch {
    return [];
  }
}

function normalizeCode(code) {
  return code.trim().replace(/\r?\n/g, "\r\n");
}

function readBasFile(methodName) {
  const direct = path.join(sourceDir, `${methodName}.bas`);
  if (fs.existsSync(direct)) {
    return normalizeCode(fs.readFileSync(direct, "utf-8"));
  }
  const manifest = readJson(manifestPath);
  const entry = manifest.find(
    (m) => m.methodName?.toLowerCase() === methodName.toLowerCase(),
  );
  if (entry?.source) {
    const p = path.join(sourceDir, entry.source);
    if (fs.existsSync(p)) return normalizeCode(fs.readFileSync(p, "utf-8"));
  }
  return null;
}

function metaFor(methodName) {
  const fromJson = readJson(jsonPath).find(
    (m) => m.methodName?.toLowerCase() === methodName.toLowerCase(),
  );
  if (fromJson) {
    return {
      methodName: fromJson.methodName,
      description: fromJson.description ?? "",
      category: fromJson.category ?? "genel",
      active: fromJson.active !== false,
      code: fromJson.code ? normalizeCode(fromJson.code) : null,
    };
  }
  const fromManifest = readJson(manifestPath).find(
    (m) => m.methodName?.toLowerCase() === methodName.toLowerCase(),
  );
  if (fromManifest) {
    return {
      methodName: fromManifest.methodName,
      description: fromManifest.description ?? "",
      category: fromManifest.category ?? "genel",
      active: fromManifest.active !== false,
      code: null,
    };
  }
  return {
    methodName,
    description: "",
    category: "genel",
    active: true,
    code: null,
  };
}

function listAllBasModules() {
  if (!fs.existsSync(sourceDir)) return [];
  return fs
    .readdirSync(sourceDir)
    .filter((f) => f.endsWith(".bas") && !SKIP_BAS.has(f))
    .map((f) => f.replace(/\.bas$/i, ""));
}

async function ensureModulesTable(sql) {
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
  await sql`ALTER TABLE modules ADD COLUMN IF NOT EXISTS run_count   INTEGER NOT NULL DEFAULT 0`;
  await sql`ALTER TABLE modules ADD COLUMN IF NOT EXISTS last_run_at TIMESTAMPTZ`;
}

async function upsertModule(sql, item) {
  const now = new Date().toISOString();
  const rows = await sql`
    INSERT INTO modules (method_name, description, category, active, code, created_at, updated_at)
    VALUES (
      ${item.methodName},
      ${item.description ?? null},
      ${item.category ?? "genel"},
      ${item.active ?? true},
      ${item.code},
      ${now},
      ${now}
    )
    ON CONFLICT (method_name) DO UPDATE SET
      description = EXCLUDED.description,
      category    = EXCLUDED.category,
      active      = EXCLUDED.active,
      code        = EXCLUDED.code,
      updated_at  = EXCLUDED.updated_at
    RETURNING method_name, updated_at
  `;
  return rows[0];
}

function resolveTargets(argv) {
  if (argv.includes("--all-json")) {
    return readJson(jsonPath)
      .filter((m) => m.methodName && m.code)
      .map((m) => m.methodName);
  }
  if (argv.includes("--all-bas")) {
    return listAllBasModules();
  }
  return argv.filter((a) => !a.startsWith("-"));
}

async function main() {
  const argv = process.argv.slice(2);
  if (argv.length === 0) {
    console.log(`Kullanım:
  node scripts/upsert-modules-db.mjs InstallCommandQueue WatchFolderServer
  node scripts/upsert-modules-db.mjs --all-json
  node scripts/upsert-modules-db.mjs --all-bas`);
    process.exit(1);
  }

  loadEnvLocal();
  if (!process.env.DATABASE_URL) {
    throw new Error("DATABASE_URL tanımlı değil.");
  }

  const sql = neon(process.env.DATABASE_URL);
  await ensureModulesTable(sql);

  const targets = resolveTargets(argv);
  if (targets.length === 0) {
    throw new Error("Upsert edilecek modül bulunamadı.");
  }

  const errors = [];
  let ok = 0;

  for (const name of targets) {
    const meta = metaFor(name);
    const code = readBasFile(name) ?? meta.code;
    if (!code) {
      errors.push(`${name}: .bas veya modules.json kodu yok`);
      continue;
    }
    try {
      const row = await upsertModule(sql, {
        methodName: meta.methodName,
        description: meta.description,
        category: meta.category,
        active: meta.active,
        code,
      });
      console.log(`✓ ${row.method_name}`);
      ok++;
    } catch (e) {
      errors.push(`${name}: ${String(e)}`);
    }
  }

  console.log(`\n${ok} modül modules tablosuna yazıldı.`);
  if (errors.length > 0) {
    console.error("Hatalar:");
    for (const err of errors) console.error(`  ✗ ${err}`);
    process.exit(1);
  }
}

main().catch((e) => {
  console.error(e.message ?? e);
  process.exit(1);
});
