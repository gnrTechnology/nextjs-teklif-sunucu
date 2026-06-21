<<<<<<< HEAD
-- Neon PostgreSQL: lisans tablosu
-- Neon SQL Editor veya psql ile bir kez çalıştırın.

CREATE TABLE IF NOT EXISTS licenses (
  mac_adresi VARCHAR(17) PRIMARY KEY,
  ip_adresi VARCHAR(45),
  firma_adi VARCHAR(255),
  user_adi VARCHAR(255),
  dosya_adi VARCHAR(255),
  proje_adi VARCHAR(255),
  proje_kisa_adresi TEXT,
  teklif_para_birimi_usd VARCHAR(50),
  teklif_para_birimi_euro VARCHAR(50),
  teklif_para_birimi_genel VARCHAR(50),
  genel_gider VARCHAR(50),
  kar VARCHAR(50),
  m31_degeri VARCHAR(50),
  veritabani_teklif JSONB,
  license VARCHAR(50) NOT NULL DEFAULT 'true',
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_licenses_firma_adi ON licenses (firma_adi);
=======
-- Neon PostgreSQL: lisans tablosu
-- Neon SQL Editor veya psql ile bir kez çalıştırın.

CREATE TABLE IF NOT EXISTS licenses (
  mac_adresi VARCHAR(17) PRIMARY KEY,
  ip_adresi VARCHAR(45),
  firma_adi VARCHAR(255),
  user_adi VARCHAR(255),
  dosya_adi VARCHAR(255),
  proje_adi VARCHAR(255),
  proje_kisa_adresi TEXT,
  teklif_para_birimi_usd VARCHAR(50),
  teklif_para_birimi_euro VARCHAR(50),
  teklif_para_birimi_genel VARCHAR(50),
  genel_gider VARCHAR(50),
  kar VARCHAR(50),
  m31_degeri VARCHAR(50),
  veritabani_teklif JSONB,
  license VARCHAR(50) NOT NULL DEFAULT 'true',
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_licenses_firma_adi ON licenses (firma_adi);
>>>>>>> 1057c5c83b00dbc2e40b55b442239b0d43d0937c
