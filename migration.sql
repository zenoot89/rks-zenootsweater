-- ═══════════════════════════════════════════════════════════
-- RKS — MULTI-TOKO MIGRATION
-- Jalankan di Supabase Dashboard → SQL Editor
-- ═══════════════════════════════════════════════════════════

-- 1. TABEL TOKO
CREATE TABLE IF NOT EXISTS toko (
  id          uuid DEFAULT gen_random_uuid() PRIMARY KEY,
  nama        text NOT NULL UNIQUE,
  deskripsi   text,
  created_at  timestamptz DEFAULT now()
);

-- 2. TAMBAH toko_id KE TABEL EXISTING
ALTER TABLE master_hpp          ADD COLUMN IF NOT EXISTS toko_id uuid REFERENCES toko(id);
ALTER TABLE history_kalkulasi   ADD COLUMN IF NOT EXISTS toko_id uuid REFERENCES toko(id);
ALTER TABLE operasional_bulanan ADD COLUMN IF NOT EXISTS toko_id uuid REFERENCES toko(id);
ALTER TABLE seller_profile      ADD COLUMN IF NOT EXISTS toko_id uuid REFERENCES toko(id);

-- 3. TABEL REKAP BULANAN (baru)
CREATE TABLE IF NOT EXISTS rekap_bulanan (
  id                uuid DEFAULT gen_random_uuid() PRIMARY KEY,
  toko_id           uuid REFERENCES toko(id) NOT NULL,
  periode           text NOT NULL,           -- format: "2026-03"
  total_pendapatan  bigint DEFAULT 0,
  total_penghasilan bigint DEFAULT 0,
  hpp               bigint DEFAULT 0,
  operasional       bigint DEFAULT 0,
  iklan             bigint DEFAULT 0,
  laba_rugi         bigint DEFAULT 0,
  npm               numeric DEFAULT 0,
  gpm               numeric DEFAULT 0,
  roas              numeric DEFAULT 0,
  total_order       int DEFAULT 0,
  aov               bigint DEFAULT 0,
  admin_ams         bigint DEFAULT 0,
  admin_fee         bigint DEFAULT 0,
  admin_layanan     bigint DEFAULT 0,
  admin_proses      bigint DEFAULT 0,
  isi_saldo         bigint DEFAULT 0,
  created_at        timestamptz DEFAULT now(),
  UNIQUE(toko_id, periode)
);

-- 4. SEED TOKO AWAL
INSERT INTO toko (nama, deskripsi) VALUES
  ('ALLEY',  'Toko Alley Knit')
ON CONFLICT (nama) DO NOTHING;

INSERT INTO toko (nama, deskripsi) VALUES
  ('ZENOOT', 'Toko Zenoot Sweater')
ON CONFLICT (nama) DO NOTHING;

-- 5. RLS (Row Level Security) — opsional, aktifkan jika butuh auth
-- ALTER TABLE toko              ENABLE ROW LEVEL SECURITY;
-- ALTER TABLE master_hpp        ENABLE ROW LEVEL SECURITY;
-- ALTER TABLE rekap_bulanan     ENABLE ROW LEVEL SECURITY;
-- ALTER TABLE operasional_bulanan ENABLE ROW LEVEL SECURITY;
