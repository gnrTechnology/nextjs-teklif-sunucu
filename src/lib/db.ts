import fs from "fs";
import path from "path";
import type { LicensePostBody, LicenseRecord } from "./types";

const DATA_DIR = path.join(process.cwd(), "data");
const LICENSES_FILE = path.join(DATA_DIR, "licenses.json");

function ensureDataDir(): void {
  if (!fs.existsSync(DATA_DIR)) {
    fs.mkdirSync(DATA_DIR, { recursive: true });
  }
}

function readAllLicenses(): LicenseRecord[] {
  ensureDataDir();
  if (!fs.existsSync(LICENSES_FILE)) {
    return [];
  }
  const raw = fs.readFileSync(LICENSES_FILE, "utf-8");
  try {
    const parsed = JSON.parse(raw) as LicenseRecord[] | { licenses: LicenseRecord[] };
    if (Array.isArray(parsed)) {
      return parsed;
    }
    if (parsed && Array.isArray(parsed.licenses)) {
      return parsed.licenses;
    }
    return [];
  } catch {
    return [];
  }
}

function writeAllLicenses(licenses: LicenseRecord[]): void {
  ensureDataDir();
  fs.writeFileSync(LICENSES_FILE, JSON.stringify(licenses, null, 2), "utf-8");
}

/** MAC adresini tutarlı anahtar formatına çevirir (AA:BB:CC:DD:EE:FF) */
export function normalizeMac(mac: string): string {
  const cleaned = mac.replace(/[^a-fA-F0-9]/g, "").toUpperCase();
  if (cleaned.length !== 12) {
    return mac.trim().toUpperCase();
  }
  return cleaned.match(/.{1,2}/g)!.join(":");
}

export function getLicenseByMac(mac: string): LicenseRecord | undefined {
  const normalized = normalizeMac(mac);
  return readAllLicenses().find(
    (item) => normalizeMac(item.macAdresi) === normalized,
  );
}

export function upsertLicense(body: LicensePostBody): LicenseRecord {
  const licenses = readAllLicenses();
  const normalizedMac = normalizeMac(body.macAdresi);
  const now = new Date().toISOString();
  const existingIndex = licenses.findIndex(
    (item) => normalizeMac(item.macAdresi) === normalizedMac,
  );

  const base: LicenseRecord =
    existingIndex >= 0
      ? { ...licenses[existingIndex] }
      : {
          macAdresi: normalizedMac,
          license: "true",
          createdAt: now,
          updatedAt: now,
        };

  const updated: LicenseRecord = {
    ...base,
    macAdresi: normalizedMac,
    ipAdresi: body.ipAdresi ?? base.ipAdresi,
    firmaAdi: body.firmaAdi ?? base.firmaAdi,
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

  if (existingIndex >= 0) {
    licenses[existingIndex] = updated;
  } else {
    licenses.push(updated);
  }

  writeAllLicenses(licenses);
  return updated;
}

export function listLicenses(): LicenseRecord[] {
  return readAllLicenses();
}

export function isLicensed(mac: string): boolean {
  const record = getLicenseByMac(mac);
  if (!record) return false;
  const value = record.license.toLowerCase();
  return value === "true" || value === "1" || value === "active" || value === "evet";
}
