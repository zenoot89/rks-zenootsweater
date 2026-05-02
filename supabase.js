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

// ════════════════════════════════════════════════════════
//  TOKO DROPDOWN — Lazada vibes
//  Inline dropdown di sidebar, bukan full-screen modal
// ════════════════════════════════════════════════════════

let _dropdownOpen = false;

function toggleTokoDropdown() {
    if (_dropdownOpen) closeTokoDropdown();
    else openTokoDropdown();
}

async function openTokoDropdown() {
    _dropdownOpen = true;
    const dd = document.getElementById('tokoDropdown');
    const chev = document.getElementById('tokoChevron');
    if (dd) dd.classList.add('open');
    if (chev) chev.classList.add('open');

    // Render list toko
    await renderTokoList();

    // Tutup kalau klik di luar
    setTimeout(() => {
        document.addEventListener('click', _outsideClickHandler);
    }, 50);
}

function closeTokoDropdown() {
    _dropdownOpen = false;
    const dd = document.getElementById('tokoDropdown');
    const chev = document.getElementById('tokoChevron');
    if (dd) dd.classList.remove('open');
    if (chev) chev.classList.remove('open');
    // Reset add form
    const addInput = document.getElementById('tokoAddInput');
    if (addInput) addInput.style.display = 'none';
    const err = document.getElementById('tokoAddErr');
    if (err) err.style.display = 'none';
    document.removeEventListener('click', _outsideClickHandler);
}

function _outsideClickHandler(e) {
    const switcher = document.getElementById('tokoSwitcher');
    if (switcher && !switcher.contains(e.target)) {
        closeTokoDropdown();
    }
}

async function renderTokoList() {
    const list = document.getElementById('tokoDropdownList');
    if (!list) return;
    list.innerHTML = '<div class="toko-dropdown-loading">Memuat...</div>';

    const tokoList = await tokoGetAll();
    const aktifId  = getAktifTokoId();

    if (!tokoList || tokoList.length === 0) {
        list.innerHTML = '<div class="toko-dropdown-loading">Belum ada toko. Tambah di bawah.</div>';
        return;
    }

    list.innerHTML = tokoList.map(t => `
        <div class="toko-item ${t.id === aktifId ? 'active' : ''}"
             onclick="pilihTokoDropdown('${t.id}','${t.nama}')">
            <div class="toko-item-icon">
                <svg width="15" height="15" viewBox="0 0 24 24" fill="none"
                     stroke="${t.id === aktifId ? 'white' : 'rgba(255,255,255,0.45)'}"
                     stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round">
                    <path d="M3 9l9-7 9 7v11a2 2 0 01-2 2H5a2 2 0 01-2-2z"/>
                    <polyline points="9 22 9 12 15 12 15 22"/>
                </svg>
            </div>
            <div class="toko-item-info">
                <div class="toko-item-nama">${t.nama}</div>
                <div class="toko-item-desc">${t.deskripsi || '—'}</div>
            </div>
            ${t.id === aktifId
                ? '<div class="toko-item-badge">Aktif</div>'
                : `<svg class="toko-item-arrow" width="13" height="13" viewBox="0 0 24 24"
                       fill="none" stroke="rgba(255,255,255,0.4)" stroke-width="2.5">
                       <polyline points="9 18 15 12 9 6"/>
                   </svg>`
            }
        </div>
    `).join('');
}

async function pilihTokoDropdown(id, nama) {
    closeTokoDropdown();
    setAktifToko(id, nama);

    // Reset semua data & UI upload box
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

    // Reset visual upload box
    ['boxIncome','boxOrder1','boxOrder2','boxAds','boxPerforma'].forEach(boxId => {
        const box = document.getElementById(boxId);
        if (!box) return;
        box.classList.remove('uploaded');
        const st = box.querySelector('.rk-upload-status');
        if (st) st.textContent = 'Belum upload';
    });
    ['fileIncome','fileOrder1','fileOrder2','fileAds','filePerforma'].forEach(fid => {
        const el = document.getElementById(fid);
        if (el) el.value = '';
    });
    ['statusIncome','statusOrder1','statusOrder2','statusAds'].forEach(sid => {
        const el = document.getElementById(sid);
        if (el) { el.textContent = ''; el.style.color = ''; }
    });
    ['rk_st_income','rk_st_order1','rk_st_order2','rk_st_ads'].forEach(sid => {
        const el = document.getElementById(sid);
        if (el) { el.textContent = '—'; el.style.color = ''; }
    });
    document.getElementById('warnUnknownAds')?.remove();

    if (typeof hppMaster !== 'undefined') hppMaster.length = 0;
    await syncHppFromSupabase();
    updateTokoUI();
    if (typeof updateRasioDashboard === 'function') updateRasioDashboard();
    if (typeof renderRekapTahunan   === 'function') renderRekapTahunan();
    if (typeof syncBiayaStatusBar   === 'function') syncBiayaStatusBar();
}

function showAddTokoInline() {
    const addInput = document.getElementById('tokoAddInput');
    const addBtn   = document.querySelector('.toko-dropdown-add');
    if (addInput) addInput.style.display = 'flex';
    if (addBtn)   addBtn.style.display   = 'none';
    setTimeout(() => document.getElementById('inlineNamaToko')?.focus(), 50);
}

async function tambahTokoInline() {
    const nama    = (document.getElementById('inlineNamaToko')?.value || '').trim().toUpperCase();
    const deskrip = (document.getElementById('inlineDeskrip')?.value  || '').trim();
    const errEl   = document.getElementById('tokoAddErr');

    if (!nama) {
        if (errEl) { errEl.textContent = 'Nama toko tidak boleh kosong'; errEl.style.display = 'block'; }
        return;
    }
    const result = await tokoCreate(nama, deskrip);
    if (!result) {
        if (errEl) { errEl.textContent = 'Gagal membuat toko. Coba lagi.'; errEl.style.display = 'block'; }
        return;
    }
    // Langsung masuk toko baru
    await pilihTokoDropdown(result.id, result.nama);
}

function updateTokoUI() {
    const nama = getAktifTokoNama();
    const badge = document.getElementById('tokoBadge');
    if (badge) badge.textContent = nama || '— Pilih Toko';
    const sw = document.querySelector('.toko-switcher-inner');
    if (sw) sw.classList.toggle('no-toko', !nama);
}

// Backward compat — beberapa tempat masih panggil showTokoModal
function showTokoModal(force = false) {
    openTokoDropdown();
}

// Init session — panggil saat app load
async function initTokoSession() {
    const savedId   = localStorage.getItem('rks_aktif_toko_id');
    const savedNama = localStorage.getItem('rks_aktif_toko_nama');

    if (savedId && savedNama) {
        const data = await supaFetch(`toko?id=eq.${savedId}&select=id,nama&limit=1`);
        if (data && data[0]) {
            setAktifToko(data[0].id, data[0].nama);
            await syncHppFromSupabase();
            updateTokoUI();
            return true;
        }
    }
    // Belum ada toko — buka dropdown
    setTimeout(() => openTokoDropdown(), 400);
    return false;
}
