# RKS — Rasio Keuangan Shopee
### Zenoot Sweater · by Burhanmology

Tool analisis keuangan untuk seller Shopee — kalkulasi margin, proyeksi harga, dan dashboard rasio keuangan bulanan.

---

## 📁 Struktur File

| File | Fungsi |
|------|--------|
| `index.html` | Shell utama — routing & layout |
| `styles.css` | Semua CSS & styling |
| `app_core.js` | Logika kalkulasi utama (4000+ baris) |
| `supabase.js` | Koneksi & CRUD ke Supabase database |

## 🗄️ Database (Supabase)

- `master_hpp` — Data HPP per SKU
- `history_kalkulasi` — Riwayat kalkulasi per order
- `operasional_bulanan` — Biaya operasional bulanan
- `seller_profile` — Profil toko

## 🚀 Deploy

Dihosting via **GitHub Pages**:  
👉 https://zenoot89.github.io/rks-zenootsweater

---

*V4.12 · conceptualized by Burhanmology*
