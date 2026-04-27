// ═══════════════════════════════════════════════════════
//  SUPABASE CONFIG — rks-zenootsweater
//  Project : rks-zenootsweater
//  URL     : https://ayhqjjzkkadrgourelqi.supabase.co
// ═══════════════════════════════════════════════════════

const SUPA_URL = 'https://ayhqjjzkkadrgourelqi.supabase.co';
const SUPA_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImF5aHFqanpra2FkcmdvdXJlbHFpIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzcyODYzODMsImV4cCI6MjA5Mjg2MjM4M30.jKKOE-ducNDAm-mt9ZnqtDjTR-JTQRZMaJTsPwUlWoM';

// ── HELPER: fetch wrapper ke Supabase REST API ──────────
async function supaFetch(endpoint, options = {}) {
    const url = `${SUPA_URL}/rest/v1/${endpoint}`;
    const headers = {
        'apikey': SUPA_KEY,
        'Authorization': `Bearer ${SUPA_KEY}`,
        'Content-Type': 'application/json',
        'Prefer': 'return=representation',
        ...options.headers
    };
    try {
        const res = await fetch(url, { ...options, headers });
        if (!res.ok) {
            const err = await res.text();
            console.error('Supabase error:', res.status, err);
            return null;
        }
        const text = await res.text();
        return text ? JSON.parse(text) : [];
    } catch (e) {
        console.error('Supabase fetch error:', e);
        return null;
    }
}

// ════════════════════════════════════════════════════════
//  MASTER HPP — CRUD
// ════════════════════════════════════════════════════════

// Ambil semua data HPP dari Supabase
async function hppGetAll() {
    const data = await supaFetch('master_hpp?select=*&order=sku.asc');
    return data || [];
}

// Simpan atau update HPP (upsert berdasarkan SKU)
async function hppUpsert(sku, nama_produk, hpp, satuan = 'pcs') {
    const payload = { sku, nama_produk, hpp: parseInt(hpp), satuan };
    const data = await supaFetch('master_hpp?on_conflict=sku', {
        method: 'POST',
        headers: { 'Prefer': 'resolution=merge-duplicates,return=representation' },
        body: JSON.stringify(payload)
    });
    return data;
}

// Hapus HPP berdasarkan SKU
async function hppDelete(sku) {
    const data = await supaFetch(`master_hpp?sku=eq.${encodeURIComponent(sku)}`, {
        method: 'DELETE'
    });
    return data;
}

// ════════════════════════════════════════════════════════
//  HISTORY KALKULASI — CRUD
// ════════════════════════════════════════════════════════

// Simpan hasil kalkulasi ke history
async function historySave(payload) {
    const data = await supaFetch('history_kalkulasi', {
        method: 'POST',
        body: JSON.stringify(payload)
    });
    return data;
}

// Ambil history (limit 50 terbaru)
async function historyGetAll() {
    const data = await supaFetch('history_kalkulasi?select=*&order=created_at.desc&limit=50');
    return data || [];
}

// Hapus history berdasarkan ID
async function historyDelete(id) {
    return await supaFetch(`history_kalkulasi?id=eq.${id}`, { method: 'DELETE' });
}

// ════════════════════════════════════════════════════════
//  OPERASIONAL BULANAN — CRUD
// ════════════════════════════════════════════════════════

// Simpan atau update operasional bulanan (upsert berdasarkan bulan)
async function opsBulananUpsert(bulan, data_ops) {
    const payload = { bulan, ...data_ops, total_ops: Object.values(data_ops).reduce((a, b) => a + (parseInt(b) || 0), 0) };
    const data = await supaFetch('operasional_bulanan?on_conflict=bulan', {
        method: 'POST',
        headers: { 'Prefer': 'resolution=merge-duplicates,return=representation' },
        body: JSON.stringify(payload)
    });
    return data;
}

// Ambil semua data operasional
async function opsGetAll() {
    const data = await supaFetch('operasional_bulanan?select=*&order=bulan.desc');
    return data || [];
}

// ════════════════════════════════════════════════════════
//  SELLER PROFILE
// ════════════════════════════════════════════════════════

async function profileGet() {
    const data = await supaFetch('seller_profile?select=*&limit=1');
    return data && data.length > 0 ? data[0] : null;
}

async function profileSave(payload) {
    const existing = await profileGet();
    if (existing) {
        return await supaFetch(`seller_profile?id=eq.${existing.id}`, {
            method: 'PATCH',
            body: JSON.stringify(payload)
        });
    } else {
        return await supaFetch('seller_profile', {
            method: 'POST',
            body: JSON.stringify(payload)
        });
    }
}

// ════════════════════════════════════════════════════════
//  SYNC: Load HPP dari Supabase ke hppMaster lokal
// ════════════════════════════════════════════════════════

async function syncHppFromSupabase() {
    const rows = await hppGetAll();
    if (!rows || rows.length === 0) return;

    // Merge ke hppMaster (array global di app_core.js)
    rows.forEach(row => {
        const existing = hppMaster.findIndex(h => h.refSku.toLowerCase() === row.sku.toLowerCase());
        if (existing >= 0) {
            hppMaster[existing] = { refSku: row.sku, namaProduk: row.nama_produk, hpp: row.hpp, satuan: row.satuan };
        } else {
            hppMaster.push({ refSku: row.sku, namaProduk: row.nama_produk, hpp: row.hpp, satuan: row.satuan });
        }
    });

    console.log(`✅ Synced ${rows.length} HPP dari Supabase`);
    renderHppTable(); // refresh tampilan Master Data
}

// ════════════════════════════════════════════════════════
//  STATUS KONEKSI
// ════════════════════════════════════════════════════════

async function checkSupabaseConnection() {
    const badge = document.getElementById('supaBadge');
    if (!badge) return;
    try {
        const data = await supaFetch('master_hpp?select=count&limit=1');
        badge.textContent = '🟢 Supabase Connected';
        badge.style.color = '#166534';
        badge.style.background = '#dcfce7';
    } catch {
        badge.textContent = '🔴 Supabase Offline';
        badge.style.color = '#991b1b';
        badge.style.background = '#fee2e2';
    }
}
