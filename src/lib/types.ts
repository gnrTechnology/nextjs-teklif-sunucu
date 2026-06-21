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
  methodName: string;
  description?: string;
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
