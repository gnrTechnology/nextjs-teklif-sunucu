export interface LicenseRecord {
  macAdresi: string;
  ipAdresi?: string;
  firmaAdi?: string;
  dosyaAdi?: string;
  projeAdi?: string;
  projeKisaAdresi?: string;
  teklifParaBirimiUSD?: string;
  teklifParaBirimiEuro?: string;
  teklifParaBirimiGenel?: string;
  genelGider?: string;
  kar?: string;
  m31Degeri?: string;
  veritabaniTeklif?: VeritabaniTeklifItem[];
  license: string;
  createdAt: string;
  updatedAt: string;
}

export interface VeritabaniTeklifItem {
  colB?: string;
  colC?: string;
  colD?: string;
  colF?: string;
  colG?: string;
  colParaBirimi?: string;
}

export interface LicensePostBody {
  macAdresi: string;
  ipAdresi?: string;
  firmaAdi?: string;
  dosyaAdi?: string;
  projeAdi?: string;
  projeKisaAdresi?: string;
  teklifParaBirimiUSD?: string;
  teklifParaBirimiEuro?: string;
  teklifParaBirimiGenel?: string;
  genelGider?: string;
  kar?: string;
  m31Degeri?: string;
  veritabaniTeklif?: VeritabaniTeklifItem[];
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
