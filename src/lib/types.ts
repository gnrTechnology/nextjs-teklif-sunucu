export interface LicenseRecord {
  macAdresi: string;
  ipAdresi?: string;
  firmaAdi?: string;
  userAdi?: string;
  dosyaAdi?: string;
  license: string;
  createdAt: string;
  updatedAt: string;
}

export interface LicensePostBody {
  macAdresi: string;
  ipAdresi?: string;
  firmaAdi?: string;
  userAdi?: string;
  dosyaAdi?: string;
}

export interface ModulePostBody {
  methodName: string;
}

export interface ModuleRecord {
  id?: number;
  methodName: string;
  description?: string;
  category?: string;
  code: string;
  active?: boolean;
  runCount?: number;
  lastRunAt?: string;
  createdAt?: string;
  updatedAt?: string;
}

export interface ModuleUpsertBody {
  methodName: string;
  description?: string;
  category?: string;
  code: string;
  active?: boolean;
}

export interface DownloadPostBody {
  macAdresi: string;
}

export interface FirmAutoStartModule {
  methodName: string;
  order: number;
  delaySeconds?: number;
  /** true ise istemci registry'de bir kez calistirildi olarak isaretler */
  runOnce?: boolean;
}

export interface FirmAutoModuleRecord {
  firmaAdi: string;
  description?: string;
  enabled?: boolean;
  isDefault?: boolean;
  onExcelOpen: {
    enabled: boolean;
    modules: FirmAutoStartModule[];
  };
}

export interface FirmAutoStartResponse {
  firmaAdi: string;
  modules: FirmAutoStartModule[];
}

export interface HeartbeatRecord {
  mac: string;
  hostname: string | null;
  userName: string | null;
  excelVersion: string | null;
  ipAddress: string | null;
  lastSeen: string;
}

export interface LicenseLog {
  id: number;
  macAdresi: string | null;
  firmaAdi: string | null;
  userAdi: string | null;
  dosyaAdi: string | null;
  ipAdresi: string | null;
  eventType: string;
  details: string | null;
  createdAt: string;
}

export type ActivityCategory =
  | "all"
  | "lisans"
  | "ihlal"
  | "guncelleme"
  | "dashboard"
  | "modul"
  | "heartbeat"
  | "komut"
  | "klasor"
  | "cihaz";

export interface UnifiedActivityItem {
  id: string;
  category: Exclude<ActivityCategory, "all">;
  title: string;
  detail?: string | null;
  mac?: string | null;
  hostname?: string | null;
  source?: string | null;
  createdAt: string;
}

export interface FolderWatchEvent {
  id: number;
  mac: string;
  hostname?: string | null;
  folderPath: string;
  eventType: "created" | "deleted" | "modified" | "started" | "scan";
  fileName?: string | null;
  filePath?: string | null;
  detail?: string | null;
  createdAt: string;
}
