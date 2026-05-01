// ═══════════════════════════════════════════════════════
//  SUPABASE CONFIG — rks-zenootsweater
//  V2.0 : Multi-Toko Support
// ═══════════════════════════════════════════════════════

const SUPA_URL = 'https://ayhqjjzkkadrgourelqi.supabase.co';
const SUPA_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImF5aHFqanpra2FkcmdvdXJlbHFpIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzcyODYzODMsImV4cCI6MjA5Mjg2MjM4M30.jKKOE-ducNDAm-mt9ZnqtDjTR-JTQRZMaJTsPwUlWoM';

// ── State toko aktif ─────────────────────────────────────
let _aktivTokoId   = localStorage.getItem('rks_aktif_toko_id')  || null;
let _aktivTokoNama = localStorage.getItem('rks_aktif_toko_nama') || null;

function getAktifTokoId()   { return _aktivTokoId; }
function getAktifTokoNama() { return _aktivTokoNama; }

function setAktifToko(id, nama) {
    _aktivTokoId   = id;
    _aktivTokoNama = nama;
    localStorage.setItem('rks_aktif_toko_id',   id);
    localStorage.setItem('rks_aktif_toko_nama', nama);
    const badge = document.getElementById('tokoBadge');
    if (badge) badge.textContent = '🏪 ' + nama;
}

// ── HELPER fetch ─────────────────────────────────────────
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
        if (!res.ok) { console.error('Supa error:', res.status, await res.text()); return null; }
        const text = await res.text();
        return text ? JSON.parse(text) : [];
    } catch (e) { console.error('Supa fetch error:', e); return null; }
}

// ════════════════════════════════════════════════════════
//  TOKO
// ════════════════════════════════════════════════════════

async function tokoGetAll() {
    return (await supaFetch('toko?select=*&order=nama.asc')) || [];
}

async function tokoCreate(nama, deskripsi = '') {
    const data = await supaFetch('toko?on_conflict=nama', {
        method: 'POST',
        headers: { 'Prefer': 'resolution=merge-duplicates,return=representation' },
        body: JSON.stringify({ nama: nama.toUpperCase().trim(), deskripsi })
    });
    return data && data[0] ? data[0] : null;
}

// ════════════════════════════════════════════════════════
//  MASTER HPP (per toko)
// ════════════════════════════════════════════════════════

async function hppGetAll() {
    const tid = getAktifTokoId(); if (!tid) return [];
    return (await supaFetch(`master_hpp?toko_id=eq.${tid}&select=*&order=sku.asc`)) || [];
}

async function hppUpsert(sku, nama_produk, hpp, variasi = '') {
    const tid = getAktifTokoId(); if (!tid) return null;
    return await supaFetch('master_hpp?on_conflict=sku,toko_id', {
        method: 'POST',
        headers: { 'Prefer': 'resolution=merge-duplicates,return=representation' },
        body: JSON.stringify({ sku, nama_produk, hpp: parseInt(hpp), variasi, toko_id: tid })
    });
}

async function hppDelete(sku) {
    const tid = getAktifTokoId(); if (!tid) return null;
    return await supaFetch(
        `master_hpp?sku=eq.${encodeURIComponent(sku)}&toko_id=eq.${tid}`,
        { method: 'DELETE' }
    );
}

// ════════════════════════════════════════════════════════
//  REKAP BULANAN (per toko)
// ════════════════════════════════════════════════════════

async function rekapSave(periode, data) {
    const tid = getAktifTokoId(); if (!tid) return null;
    return await supaFetch('rekap_bulanan?on_conflict=toko_id,periode', {
        method: 'POST',
        headers: { 'Prefer': 'resolution=merge-duplicates,return=representation' },
        body: JSON.stringify({ toko_id: tid, periode, ...data })
    });
}

async function rekapGetAll() {
    const tid = getAktifTokoId(); if (!tid) return [];
    return (await supaFetch(
        `rekap_bulanan?toko_id=eq.${tid}&select=*&order=periode.asc`
    )) || [];
}

async function rekapDelete(periode) {
    const tid = getAktifTokoId(); if (!tid) return null;
    return await supaFetch(
        `rekap_bulanan?toko_id=eq.${tid}&periode=eq.${periode}`,
        { method: 'DELETE' }
    );
}

// Semua toko untuk perbandingan di Rekap Tahunan
async function rekapGetAllToko() {
    return (await supaFetch(
        'rekap_bulanan?select=*,toko(nama)&order=toko_id.asc,periode.asc'
    )) || [];
}

// ════════════════════════════════════════════════════════
//  HISTORY KALKULASI (per toko)
// ════════════════════════════════════════════════════════

async function historySave(payload) {
    const tid = getAktifTokoId(); if (!tid) return null;
    return await supaFetch('history_kalkulasi', {
        method: 'POST',
        body: JSON.stringify({ ...payload, toko_id: tid })
    });
}

async function historyGetAll() {
    const tid = getAktifTokoId(); if (!tid) return [];
    return (await supaFetch(
        `history_kalkulasi?toko_id=eq.${tid}&select=*&order=created_at.desc&limit=50`
    )) || [];
}

async function historyDelete(id) {
    return await supaFetch(`history_kalkulasi?id=eq.${id}`, { method: 'DELETE' });
}

// ════════════════════════════════════════════════════════
//  OPERASIONAL BULANAN (per toko)
// ════════════════════════════════════════════════════════

async function opsBulananUpsert(bulan, data_ops) {
    const tid = getAktifTokoId(); if (!tid) return null;
    const total_ops = Object.values(data_ops).reduce((a,b) => a + (parseInt(b)||0), 0);
    return await supaFetch('operasional_bulanan?on_conflict=bulan,toko_id', {
        method: 'POST',
        headers: { 'Prefer': 'resolution=merge-duplicates,return=representation' },
        body: JSON.stringify({ bulan, toko_id: tid, ...data_ops, total_ops })
    });
}

async function opsGetAll() {
    const tid = getAktifTokoId(); if (!tid) return [];
    return (await supaFetch(
        `operasional_bulanan?toko_id=eq.${tid}&select=*&order=bulan.desc`
    )) || [];
}

// ════════════════════════════════════════════════════════
//  SELLER PROFILE (per toko)
// ════════════════════════════════════════════════════════

async function profileGet() {
    const tid = getAktifTokoId(); if (!tid) return null;
    const data = await supaFetch(`seller_profile?toko_id=eq.${tid}&select=*&limit=1`);
    return data && data[0] ? data[0] : null;
}

async function profileSave(payload) {
    const tid = getAktifTokoId(); if (!tid) return null;
    const existing = await profileGet();
    if (existing) {
        return await supaFetch(`seller_profile?id=eq.${existing.id}`, {
            method: 'PATCH',
            body: JSON.stringify({ ...payload, toko_id: tid })
        });
    }
    return await supaFetch('seller_profile', {
        method: 'POST',
        body: JSON.stringify({ ...payload, toko_id: tid })
    });
}

// ════════════════════════════════════════════════════════
//  SYNC HPP Supabase → hppMaster lokal
// ════════════════════════════════════════════════════════

async function syncHppFromSupabase() {
    const rows = await hppGetAll();
    if (!rows || rows.length === 0) return;
    if (typeof hppMaster === 'undefined') return;
    hppMaster.length = 0;
    rows.forEach(row => {
        hppMaster.push({
            refSku:     row.sku,
            namaProduk: row.nama_produk,
            variasi:    row.variasi || '',
            hpp:        row.hpp
        });
    });
    if (typeof renderHppTable === 'function') renderHppTable();
    console.log(`✅ ${rows.length} HPP synced (${getAktifTokoNama()})`);
}

// ════════════════════════════════════════════════════════
//  INIT SESSION TOKO — panggil saat app load
// ════════════════════════════════════════════════════════

async function initTokoSession() {
    const savedId   = localStorage.getItem('rks_aktif_toko_id');
    const savedNama = localStorage.getItem('rks_aktif_toko_nama');

    if (savedId && savedNama) {
        // Verifikasi ke Supabase
        const data = await supaFetch(`toko?id=eq.${savedId}&select=id,nama&limit=1`);
        if (data && data[0]) {
            setAktifToko(data[0].id, data[0].nama);
            await syncHppFromSupabase();
            updateTokoUI();
            return true;
        }
    }
    // Belum ada → paksa modal
    await showTokoModal(true);
    return false;
}

// ════════════════════════════════════════════════════════
//  MODAL PILIH / TAMBAH TOKO
// ════════════════════════════════════════════════════════

async function showTokoModal(force = false) {
    const old = document.getElementById('tokoModal');
    if (old) old.remove();

    const tokoList = await tokoGetAll();

    const modal = document.createElement('div');
    modal.id = 'tokoModal';
    modal.style.cssText = `
        position:fixed;inset:0;background:rgba(0,0,0,0.75);backdrop-filter:blur(4px);
        display:flex;align-items:center;justify-content:center;z-index:99999;
    `;

    modal.innerHTML = `
        <div style="background:#1a1d24;border:1px solid #2a2d35;border-radius:16px;
                    padding:32px;width:440px;max-width:92vw;box-shadow:0 24px 64px rgba(0,0,0,0.6)">

            <div style="display:flex;align-items:center;gap:12px;margin-bottom:8px">
                <div style="width:44px;height:44px;background:#ef4444;border-radius:12px;
                            display:flex;align-items:center;justify-content:center;font-size:22px">🏪</div>
                <div>
                    <div style="font-size:17px;font-weight:700;color:#f1f5f9">Pilih Toko Aktif</div>
                    <div style="font-size:12px;color:#6b7280">Data HPP & Rekap tersimpan terpisah per toko</div>
                </div>
            </div>

            <div style="border-bottom:1px solid #2a2d35;margin:16px 0"></div>

            <!-- List toko -->
            <div style="display:flex;flex-direction:column;gap:8px;margin-bottom:20px;max-height:240px;overflow-y:auto">
                ${tokoList.length === 0
                    ? `<div style="color:#6b7280;font-size:13px;text-align:center;padding:20px">
                         Belum ada toko. Buat toko pertama di bawah.
                       </div>`
                    : tokoList.map(t => `
                        <button onclick="pilihToko('${t.id}','${t.nama}')"
                            style="background:#0e0f11;border:1px solid #2a2d35;border-radius:10px;
                                   padding:14px 16px;text-align:left;cursor:pointer;color:#f1f5f9;
                                   display:flex;align-items:center;justify-content:space-between;
                                   transition:all 0.15s;width:100%"
                            onmouseover="this.style.borderColor='#ef4444';this.style.background='#1a0808'"
                            onmouseout="this.style.borderColor='#2a2d35';this.style.background='#0e0f11'">
                            <div>
                                <div style="font-size:14px;font-weight:700">🏪 ${t.nama}</div>
                                <div style="font-size:11px;color:#6b7280;margin-top:2px">${t.deskripsi||'—'}</div>
                            </div>
                            <span style="color:#ef4444;font-size:13px;font-weight:600">Masuk →</span>
                        </button>
                    `).join('')
                }
            </div>

            <!-- Tambah toko baru -->
            <div style="background:#0e0f11;border:1px solid #2a2d35;border-radius:10px;padding:16px">
                <div style="font-size:11px;font-weight:700;color:#6b7280;
                            text-transform:uppercase;letter-spacing:1px;margin-bottom:10px">
                    ＋ Tambah Toko Baru
                </div>
                <div style="display:flex;gap:8px;margin-bottom:8px">
                    <input id="inputNamaToko" type="text" placeholder="Nama (cth: ALLEY)"
                        maxlength="30"
                        style="flex:1;background:#1a1d24;border:1px solid #2a2d35;border-radius:8px;
                               color:#f1f5f9;padding:9px 11px;font-size:13px;outline:none;
                               text-transform:uppercase;font-family:inherit"
                        oninput="this.value=this.value.toUpperCase()"
                        onkeydown="if(event.key==='Enter')tambahToko()">
                    <input id="inputDeskrip" type="text" placeholder="Deskripsi (opsional)"
                        style="flex:1.6;background:#1a1d24;border:1px solid #2a2d35;border-radius:8px;
                               color:#f1f5f9;padding:9px 11px;font-size:13px;outline:none;font-family:inherit"
                        onkeydown="if(event.key==='Enter')tambahToko()">
                    <button onclick="tambahToko()"
                        style="background:#ef4444;color:#fff;border:none;border-radius:8px;
                               padding:9px 16px;cursor:pointer;font-weight:700;font-size:13px;
                               white-space:nowrap">Buat</button>
                </div>
                <div id="tokoErrMsg" style="color:#f87171;font-size:11px;display:none"></div>
            </div>

            ${!force && _aktivTokoId ? `
            <div style="margin-top:14px;text-align:center">
                <button onclick="document.getElementById('tokoModal').remove()"
                    style="background:transparent;border:none;color:#6b7280;
                           font-size:12px;cursor:pointer;text-decoration:underline">
                    Batal — tetap di toko ${_aktivTokoNama||''}
                </button>
            </div>` : ''}
        </div>
    `;
    document.body.appendChild(modal);
    setTimeout(() => document.getElementById('inputNamaToko')?.focus(), 100);
}

async function pilihToko(id, nama) {
    setAktifToko(id, nama);
    document.getElementById('tokoModal')?.remove();

    // ── Reset semua data & UI upload box ──────────────────────
    if (typeof rkData !== 'undefined') {
        rkData.income           = null;
        rkData.order1           = null;
        rkData.order2           = null;
        rkData.ads              = null;
        rkData.performa         = null;
        rkData.periode          = '';
        rkData._incomeNoPesanan = new Set();
        rkData._incomeRawRows   = [];
        rkData._orderRawMap     = {};
    }

    // ── Reset visual upload box ke state awal ─────────────────
    ['boxIncome','boxOrder1','boxOrder2','boxAds','boxPerforma'].forEach(boxId => {
        const box = document.getElementById(boxId);
        if (!box) return;
        box.classList.remove('uploaded');
        const st = box.querySelector('.rk-upload-status');
        if (st) st.textContent = 'Belum upload';
    });
    // Reset input file agar bisa re-upload file yang sama
    ['fileIncome','fileOrder1','fileOrder2','fileAds','filePerforma'].forEach(fid => {
        const el = document.getElementById(fid);
        if (el) el.value = '';
    });
    // Reset status text checklist bawah
    ['statusIncome','statusOrder1','statusOrder2','statusAds'].forEach(sid => {
        const el = document.getElementById(sid);
        if (el) { el.textContent = ''; el.style.color = ''; }
    });
    ['rk_st_income','rk_st_order1','rk_st_order2','rk_st_ads'].forEach(sid => {
        const el = document.getElementById(sid);
        if (el) { el.textContent = '—'; el.style.color = ''; }
    });
    // Hapus warning unknown ads
    document.getElementById('warnUnknownAds')?.remove();

    if (typeof hppMaster !== 'undefined') hppMaster.length = 0;
    await syncHppFromSupabase();
    updateTokoUI();
    if (typeof updateRasioDashboard === 'function') updateRasioDashboard();
    if (typeof renderRekapTahunan   === 'function') renderRekapTahunan();
    if (typeof syncBiayaStatusBar   === 'function') syncBiayaStatusBar();
}

async function tambahToko() {
    const nama    = (document.getElementById('inputNamaToko')?.value||'').trim().toUpperCase();
    const deskrip = (document.getElementById('inputDeskrip')?.value||'').trim();
    const errEl   = document.getElementById('tokoErrMsg');
    if (!nama) {
        errEl.textContent = 'Nama toko tidak boleh kosong'; errEl.style.display='block'; return;
    }
    const result = await tokoCreate(nama, deskrip);
    if (!result) {
        errEl.textContent = 'Gagal membuat toko. Coba lagi.'; errEl.style.display='block'; return;
    }
    await pilihToko(result.id, result.nama);
}

function updateTokoUI() {
    const nama = getAktifTokoNama();
    const badge = document.getElementById('tokoBadge');
    if (badge) badge.textContent = nama ? nama : '— Pilih Toko';
    const sw = document.querySelector('.toko-switcher-inner');
    if (sw) sw.classList.toggle('no-toko', !nama);
}

// ════════════════════════════════════════════════════════
//  CEK STATUS KONEKSI
// ════════════════════════════════════════════════════════

async function checkSupabaseConnection() {
    const badge = document.getElementById('supaBadge');
    if (!badge) return;
    try {
        const d = await supaFetch('toko?select=id&limit=1');
        badge.textContent = '🟢 Supabase OK';
        badge.style.color = '#166534'; badge.style.background = '#dcfce7';
    } catch {
        badge.textContent = '🔴 Offline';
        badge.style.color = '#991b1b'; badge.style.background = '#fee2e2';
    }
}
