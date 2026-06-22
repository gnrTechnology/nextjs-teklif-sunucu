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
