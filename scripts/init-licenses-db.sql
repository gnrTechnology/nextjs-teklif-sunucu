-- Neon PostgreSQL: lisans tablosu
-- Neon SQL Editor'da çalıştırın.

CREATE TABLE IF NOT EXISTS licenses (
  mac_adresi VARCHAR(17) PRIMARY KEY,
  ip_adresi  VARCHAR(45),
  firma_adi  VARCHAR(255),
  user_adi   VARCHAR(255),
  dosya_adi  VARCHAR(255),
  license    VARCHAR(50) NOT NULL DEFAULT 'true',
  created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- Mevcut tabloda fazla kolonlar varsa temizle
ALTER TABLE licenses DROP COLUMN IF EXISTS proje_adi;
ALTER TABLE licenses DROP COLUMN IF EXISTS proje_kisa_adresi;
ALTER TABLE licenses DROP COLUMN IF EXISTS teklif_para_birimi_usd;
ALTER TABLE licenses DROP COLUMN IF EXISTS teklif_para_birimi_euro;
ALTER TABLE licenses DROP COLUMN IF EXISTS teklif_para_birimi_genel;
ALTER TABLE licenses DROP COLUMN IF EXISTS genel_gider;
ALTER TABLE licenses DROP COLUMN IF EXISTS kar;
ALTER TABLE licenses DROP COLUMN IF EXISTS m31_degeri;
ALTER TABLE licenses DROP COLUMN IF EXISTS veritabani_teklif;

-- user_adi yoksa ekle
ALTER TABLE licenses ADD COLUMN IF NOT EXISTS user_adi VARCHAR(255);

CREATE INDEX IF NOT EXISTS idx_licenses_firma_adi ON licenses (firma_adi);
