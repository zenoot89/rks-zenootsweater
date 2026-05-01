
let d = { harga:0,voucherToko:0,subsidiShopee:0,subtotal:0,cair:0,admin:0,layanan:0,proses:0,aff:0,kampanye:0,saldo:0,qty:1,existsAff:false,existsKampanye:false };

const formatRp = v => new Intl.NumberFormat('id-ID',{style:'currency',currency:'IDR',minimumFractionDigits:0,maximumFractionDigits:0}).format(Math.floor(v));

// Tab-tab yang TIDAK menampilkan panel Rekomendasi Aksi (kanan)
const TABS_NO_REKOMEN = new Set(['rasio','rasio_upload','rasio_rekap','masterdata','history','analisis_produk','analisis_biaya']);

function switchTab(name,btn){
    document.querySelectorAll('.tab-panel').forEach(p=>p.classList.remove('active'));
    document.querySelectorAll('.tab-btn').forEach(b=>b.classList.remove('active'));
    document.getElementById('tab-'+name).classList.add('active');
    // Activate all matching sidebar buttons
    document.querySelectorAll('.tab-btn').forEach(b=>{
        if(b.getAttribute('onclick') && b.getAttribute('onclick').includes("'"+name+"'")) b.classList.add('active');
    });
    // Sembunyikan panel kanan untuk tab Rasio Keuangan & Lainnya
    const rp = document.getElementById('rightPanel');
    if (rp) rp.style.display = TABS_NO_REKOMEN.has(name) ? 'none' : '';
    if(name==='simulator') updateSimulator();
    if(name==='history') renderHistory();
    if(name==='reverse') hitungReverse();
    if(name==='analisis_produk') tryAutoLoadSkuData();
    if(name==='analisis_biaya') syncBiayaStatusBar();
}

function formatInputRibuan(el){
    let raw=el.value.replace(/\./g,'').replace(/[^0-9]/g,'');
    if(!raw) return;
    el.value=parseInt(raw).toLocaleString('id-ID');
}

function smartParseNumber(str) {
    // Hapus semua karakter kecuali angka, titik, koma
    let s = String(str).replace(/[^0-9.,]/g, '').trim();
    if (!s) return 0;

    const hasDot   = s.includes('.');
    const hasComma = s.includes(',');

    if (hasDot && hasComma) {
        // Kedua ada: yang terakhir muncul = desimal
        const lastDot   = s.lastIndexOf('.');
        const lastComma = s.lastIndexOf(',');
        if (lastComma > lastDot) {
            // format: 1.000,50 (Indonesia)
            s = s.replace(/\./g, '').replace(',', '.');
        } else {
            // format: 1,000.50 (Internasional)
            s = s.replace(/,/g, '');
        }
    } else if (hasDot && !hasComma) {
        // Hanya titik — cek apakah desimal atau ribuan
        const parts = s.split('.');
        const lastPart = parts[parts.length - 1];
        if (parts.length === 2 && lastPart.length !== 3) {
            // "7.9" atau "3.14" → desimal
            // biarkan as-is
        } else {
            // "1.000" atau "1.000.000" → ribuan, hapus titik
            s = s.replace(/\./g, '');
        }
    } else if (hasComma && !hasDot) {
        // Hanya koma — cek apakah desimal atau ribuan
        const parts = s.split(',');
        const lastPart = parts[parts.length - 1];
        if (parts.length === 2 && lastPart.length !== 3) {
            // "7,9" atau "3,14" → desimal
            s = s.replace(',', '.');
        } else {
            // "1,000" atau "1,000,000" → ribuan, hapus koma
            s = s.replace(/,/g, '');
        }
    }

    return parseFloat(s) || 0;
}

function parseVal(id){
    return smartParseNumber(document.getElementById(id).value);
}

function syncNpm(source) {
    let ns = d.subtotal * (d.qty || 1);
    if (ns <= 0) return; 
    if (source === 'pct') {
        let pct = parseFloat(document.getElementById('targetNpm').value) || 0;
        let rp = Math.round(ns * (pct / 100));
        document.getElementById('targetLabaRp').value = rp.toLocaleString('id-ID');
    } else if (source === 'rp') {
        let rp = parseVal('targetLabaRp');
        let pct = (rp / ns) * 100;
        document.getElementById('targetNpm').value = pct.toFixed(2);
    }
}

function syncAds(source) {
    if (source === 'roas') {
        let roas = smartParseNumber(document.getElementById('roas').value);
        let acos = roas > 0 ? (100 / roas).toFixed(1) : 0;
        document.getElementById('acos').value = acos;
    } else if (source === 'acos') {
        let acos = smartParseNumber(document.getElementById('acos').value);
        let roas = acos > 0 ? (100 / acos).toFixed(2) : 0;
        document.getElementById('roas').value = roas;
    }
}

function cleanAndHitung(){simpanPermanen();hitung();}

function simpanPermanen(){
    ['oprTotalMonth','oprPctTarget','hppSatuan','roas','acos','targetNpm','pctAffManual','maxOrderBln'].forEach(id=>{
        localStorage.setItem(id,document.getElementById(id).value);
    });
}

function muatDataTersimpan(){
    const idsRibuan=['oprTotalMonth','hppSatuan','maxOrderBln'];
    ['oprTotalMonth','oprPctTarget','hppSatuan','roas','acos','targetNpm','pctAffManual','maxOrderBln'].forEach(id=>{
        let v=localStorage.getItem(id);
        if(v){
            document.getElementById(id).value=v;
            // Format ulang input yang pakai pemisah ribuan
            if(idsRibuan.includes(id)) formatInputRibuan(document.getElementById(id));
        }
    });
    hitung(); renderHistory(); hitungReverse(); muatMasterData();
}

function hapusHanyaDataPaste(){
    document.getElementById('pasteArea1').value='';
    document.getElementById('pasteArea2').value='';
    d={harga:0,voucherToko:0,subsidiShopee:0,subtotal:0,cair:0,admin:0,layanan:0,proses:0,aff:0,kampanye:0,saldo:0,qty:1,existsAff:false,existsKampanye:false};
    document.getElementById('infoQty').innerText=1;
    sembunyikanBarisDinamis();
    document.getElementById('targetLabaRp').value='';
    document.getElementById('alertBanner').style.display='none';
    hitung();
}

function sembunyikanBarisDinamis(){['rowKampanye','rowAffiliate','rowSaldo','rowSaldoNote','rowSubtotalBasis'].forEach(id=>document.getElementById(id).classList.remove('visible'));}
function tampilkanBaris(id){document.getElementById(id).classList.add('visible');}

const createScanner=text=>keys=>{
    for(let k of keys){
        let sk=k.replace(/[.*+?^${}()|[\]\\]/g,'\\$&');
        let m=(text.match(new RegExp(sk+'[^\\d\\n]*[\\n\\s\\:]+(?:-\\s*)?(?:Rp\\s*)?(?:-\\s*)?(\\d[\\d\\.,]*)','i'))
            ||text.match(new RegExp(sk+'[^\\d\\n]*?(?:-\\s*)?(?:Rp\\s*)?(?:-\\s*)?(\\d[\\d\\.,]*)','i')));
        if(m) return Math.abs(parseFloat(m[1].replace(/[.,]/g,'')));
    }
    return 0;
};

function bacaPaste(){
    const t1=document.getElementById('pasteArea1').value;
    const t2=document.getElementById('pasteArea2').value;
    const s1=createScanner(t1), s2=createScanner(t2);
    let qty=1;
    const batas=t2.search(/Sembunyikan/i);
    const tbl=batas>0?t2.substring(0,batas):'';
    if(tbl){
        const re=/(\d{1,3}(?:\.\d{3})+|\d+)\s*\n\s*(\d{1,3})\s*\n\s*(\d{1,3}(?:\.\d{3})+|\d+)/g;
        let tot=0,m;
        while((m=re.exec(tbl))!==null){
            const q=parseInt(m[2]),h=parseFloat(m[1].replace(/\./g,'')),st=parseFloat(m[3].replace(/\./g,''));
            if(q>=1&&q<=999&&Math.abs(h*q-st)<10) tot+=q;
        }
        if(tot>0) qty=tot;
    }
    if(qty===1&&batas>0){let m=t2.match(/(\d[\d.,]*)\s*\n\s*(\d+)\s*\n\s*(\d[\d.,]*)\s*\n?\s*Sembunyikan/i);if(m) qty=parseInt(m[2])||1;}
    if(qty===1){let m=(t1+'\n'+t2).match(/x(\d+)/i)||(t1+'\n'+t2).match(/(\d+)\s*pc/i);if(m) qty=parseInt(m[1])||1;}
    d.qty=qty; document.getElementById('infoQty').innerText=qty;

    // ── AUTO LOOKUP HPP dari Master Data ──────────────
    // Coba berbagai label yang mungkin muncul di teks Shopee
    const gabung = t1+'\n'+t2;
    const refPatterns = [
        /Nomor\s+Referensi\s+SKU\s*[:\-]?\s*([A-Za-z0-9_\-]+)/i,
        /No\.?\s*Ref(?:erensi)?\s*SKU\s*[:\-]?\s*([A-Za-z0-9_\-]+)/i,
        /Kode\s+Variasi\s*[:\-]?\s*([A-Za-z0-9_\-]+)/i,
        /SKU\s+Variasi\s*[:\-]?\s*([A-Za-z0-9_\-]+)/i,
        /Referensi\s+SKU\s*[:\-]?\s*([A-Za-z0-9_\-]+)/i,
    ];
    let autoRefSku = null;
    for (const pat of refPatterns) {
        const m = gabung.match(pat);
        if (m) { autoRefSku = m[1].trim(); break; }
    }
    // Jika tidak ketemu dari label, coba cari langsung nilai yang cocok dengan master
    if (!autoRefSku && hppMaster.length > 0) {
        for (const h of hppMaster) {
            if (gabung.includes(h.refSku)) { autoRefSku = h.refSku; break; }
        }
    }
    if (autoRefSku) {
        const foundEntry = hppMaster.find(h => h.refSku.toLowerCase().includes(autoRefSku.toLowerCase()));
        if (foundEntry) {
            document.getElementById('hppSatuan').value = foundEntry.hpp.toLocaleString('id-ID');
            document.getElementById('hppRefSku').value = foundEntry.refSku;
            const el = document.getElementById('hppSatuan');
            el.style.borderColor='#16a34a'; el.style.background='#f0fdf4';
            setTimeout(()=>{ el.style.borderColor=''; el.style.background=''; }, 2500);
            // Tampilkan label SKU aktif
            const info = document.getElementById('hppRefInfo');
            const skuActive = document.getElementById('hppRefSkuActive');
            if (info) { info.style.display='block'; info.style.color='#16a34a'; info.style.fontSize='0.75em'; info.style.fontWeight='600'; info.innerHTML='✓ Approved'; }
            if (skuActive) { skuActive.style.display='none'; }
        }
    }
    
    d.subsidiShopee=s1(['Voucher Shopee','Diskon dari Shopee']);
    d.voucherToko=Math.round(s2(['Voucher Toko yang ditanggung Penjual','Voucher Penjual','Diskon dari Penjual','Voucher Toko'])/qty);
    const sp=s2(['Subtotal Pesanan','Harga Produk','Subtotal Produk']);
    d.harga=qty>0?Math.round(sp/qty):sp; 
    // Subtotal Produk = harga pasang SETELAH diskon toko (sudah embedded)
    // Ini yang jadi BASIS RASIO semua potongan Shopee — JANGAN dikurangi voucherToko lagi
    d.subtotal=d.harga;
    
    d.admin=s2(['Biaya Administrasi']); d.layanan=s2(['Biaya Layanan']); d.proses=s2(['Biaya Proses Pesanan']);
    d.aff=s2(['Biaya Komisi AMS','Komisi Affiliate','Potongan Affiliate']);
    d.kampanye=s2(['Biaya Kampanye']); d.saldo=s2(['Biaya Isi Saldo Otomatis']);
    
    d.existsAff=d.aff>0; d.existsKampanye=d.kampanye>0;
    d.cair=s2(['Total Penghasilan','Estimasi Total Penghasilan','Penghasilan Bersih']);
    
    if(!d.subsidiShopee) d.subsidiShopee=s2(['Voucher Shopee','Diskon dari Shopee']);
    
    const th=d.harga*qty, tv=d.voucherToko*qty;
    // ns = Net Sales basis = Subtotal Produk (harga setelah diskon toko, sebelum potongan Shopee)
    const ns=th;
    
    document.getElementById('infoHarga').innerText=formatRp(th);
    document.getElementById('infoVoucher').innerText='- '+formatRp(tv);
    document.getElementById('pctVoucher').innerText=th>0?((tv/th)*100).toFixed(1)+'%':'0%';
    document.getElementById('infoDiskonShopee').innerText=formatRp(d.subsidiShopee);
    document.getElementById('pctSubsidi').innerText=ns>0?((d.subsidiShopee/ns)*100).toFixed(1)+'%':'0%';
    document.getElementById('infoSubtotal').innerText=formatRp(ns);
    
    document.getElementById('infoAdmin').innerText='- '+formatRp(d.admin);
    document.getElementById('infoLayanan').innerText='- '+formatRp(d.layanan);
    document.getElementById('infoProses').innerText='- '+formatRp(d.proses);
    document.getElementById('infoCair').innerText=formatRp(d.cair);
    
    const sp2=(id,v)=>document.getElementById(id).innerText=(ns>0?((v/ns)*100).toFixed(1):0)+'%';
    sp2('pctAdmin',d.admin); sp2('pctLayanan',d.layanan); sp2('pctProses',d.proses);
    
    sembunyikanBarisDinamis();
    if(d.existsKampanye){tampilkanBaris('rowKampanye');document.getElementById('infoKampanye').innerText='- '+formatRp(d.kampanye);sp2('pctKampanye',d.kampanye);}
    if(d.existsAff){tampilkanBaris('rowAffiliate');document.getElementById('infoAffiliate').innerText='- '+formatRp(d.aff);sp2('pctAffiliate',d.aff);}
    if(d.saldo>0){
        tampilkanBaris('rowSaldo');
        tampilkanBaris('rowSaldoNote');
        document.getElementById('infoSaldo').innerText=formatRp(d.saldo);
        sp2('pctSaldo',d.saldo);
    }
    
    if(d.voucherToko > 0) tampilkanBaris('rowSubtotalBasis');
    syncNpm('pct');
    hitung();
}

function hitung(){
    const hpp=parseVal('hppSatuan'), qty=d.qty||1, roas=parseVal('roas'), pAff=parseVal('pctAffManual');
    const acosP=smartParseNumber(document.getElementById('acos').value);
    const oprI=parseVal('oprTotalMonth'), oprP=parseVal('oprPctTarget');
    const tnpm=smartParseNumber(document.getElementById('targetNpm').value);
    const maxOrderBln=parseVal('maxOrderBln');

    const tHpp=hpp*qty, ns=d.subtotal*qty;

    const ads=ns>0?Math.floor(ns*(acosP/100)):0;
    const targetOmset=oprP>0?oprI/(oprP/100):0;
    const targetOrder=ns>0?Math.ceil(targetOmset/ns):0;
    const opr=targetOrder>0?Math.floor(oprI/targetOrder):Math.floor(ns*(oprP/100));
    const fAff=d.existsAff?d.aff:Math.floor(ns*(pAff/100));

    // d.saldo = Isi Saldo Otomatis Iklan — bukan biaya riil, hanya transfer ke saldo iklan.
    // d.cair dari Shopee sudah dipotong d.saldo, jadi harus ditambahkan kembali agar laba akurat.
    const cairBersih = d.cair + d.saldo;

    const laba = d.existsAff
        ? cairBersih - (tHpp + opr + ads)
        : cairBersih - (tHpp + opr + ads + fAff);

    const npm=ns>0?(laba/ns)*100:0;
    const gpm=ns>0?((ns-tHpp)/ns)*100:0;
    const labaMonth=targetOrder*laba;
    const cashflow=targetOrder*tHpp+(roas>0?targetOmset/roas*1.11:0)+oprI;

    // Affiliate Rp live (poin 6 & 9)
    const affRp = d.existsAff ? d.aff : fAff;
    document.getElementById('affRpLive').value = formatRp(affRp);

    document.getElementById('resTargetOmset').innerText=formatRp(targetOmset);
    document.getElementById('resTargetOrderMonth').innerText=targetOrder;
    document.getElementById('resTargetOrderDay').innerText=Math.ceil(targetOrder/30);
    document.getElementById('resTotalCashflow').innerText=formatRp(cashflow);
    document.getElementById('resLabaMonth').innerText=formatRp(labaMonth);
    document.getElementById('resOprHighlight').innerText=formatRp(oprI);
    document.getElementById('resTotalIncome').innerText=formatRp(labaMonth + oprI);

    // Warna card sesuai kondisi rugi/untung
    const isRugi = labaMonth < 0;
    const cardGaji = document.getElementById('cardGaji');
    const cardIncome = document.getElementById('cardTotalIncome');
    const labelIncome = document.getElementById('labelTotalIncome');
    const subIncome = document.getElementById('subTotalIncome');

    if (labelIncome) {
        labelIncome.innerText = isRugi
            ? 'Kamu boncos per bulan'
            : 'Kamu punya duit';
    }
    if (subIncome) {
        subIncome.innerText = isRugi
            ? 'Kerugian Gaji + Operasional'
            : 'Gaji + Operasional';
    }
    if (cardGaji) {
        cardGaji.style.borderLeftColor = isRugi ? '#991b1b' : '#166534';
        cardGaji.querySelector('strong').style.color = isRugi ? '#991b1b' : '#0f172a';
    }
    if (cardIncome) {
        cardIncome.style.background    = isRugi ? '#fee2e2' : '#dcfce7';
        cardIncome.style.borderColor   = isRugi ? '#fecaca' : '#bbf7d0';
        cardIncome.style.borderLeftColor = isRugi ? '#991b1b' : '#15803d';
        cardIncome.querySelectorAll('span').forEach(s => { if(s.id!=='subTotalIncome'||true) s.style.color = isRugi ? '#991b1b' : '#166534'; });
        cardIncome.querySelector('strong').style.color = isRugi ? '#7f1d1d' : '#14532d';
    }
    // cardOpr warna selalu kuning (target tetap, bukan rugi/untung)
    const cardOprEl = document.getElementById('cardOpr');
    if (cardOprEl) {
        cardOprEl.style.background = '#fffbf0';
        cardOprEl.style.borderLeftColor = '#f59e0b';
    }

    // TOTAL ADMIN = semua potongan platform + affiliate (declared early - used multiple times)
    const totalPlatformBiaya=d.admin+d.layanan+d.proses+d.kampanye;
    const totalAdminInkAff=totalPlatformBiaya + (d.existsAff ? d.aff : Math.floor(ns*(pAff/100)));
    const aPct=ns>0?(totalPlatformBiaya/ns):0.138;
    const affPct=d.existsAff&&ns>0?(d.aff/ns):(pAff/100);

    document.getElementById('dashGpm').innerText=gpm.toFixed(1)+'%';
    document.getElementById('dashNpm').innerText=npm.toFixed(1)+'%';
    document.getElementById('dashLaba').innerText=formatRp(laba);
    document.getElementById('dashAcos').innerText=ns>0?((ads/ns)*100).toFixed(1)+'%':'0%';
    document.getElementById('dashCpp').innerText=formatRp(ads);

    // Admin% dan Gap untuk GPM card
    const adminPctDash = ns>0?(totalAdminInkAff/ns)*100:0;
    const gapPct = gpm - adminPctDash;
    const gapEl = document.getElementById('dashGap');
    const adminEl = document.getElementById('dashAdminPct');
    if(adminEl) adminEl.innerText = adminPctDash.toFixed(1)+'%';
    if(gapEl){
        gapEl.innerText = (gapPct>=0?'+':'')+gapPct.toFixed(1)+'%';
        gapEl.style.color = gapPct < 0 ? '#fca5a5' : gapPct < 15 ? '#fde68a' : '#86efac';
    }

    document.getElementById('labelOprDesc').innerText=`BEBAN OPR (${oprP}%)`;
    document.getElementById('infoOprValue').innerText='- '+formatRp(opr);

    const lb=document.getElementById('dashLabaBox');
    lb.classList.toggle('bg-red',laba<0);
    lb.classList.toggle('bg-teal',laba>=0);

    if(!d.existsAff){
        if(pAff>0){
            tampilkanBaris('rowAffiliate');
            document.getElementById('infoAffiliate').innerText='- '+formatRp(fAff);
            document.getElementById('pctAffiliate').innerText=(ns>0?((fAff/ns)*100).toFixed(1):0)+'%';
        } else {
            document.getElementById('rowAffiliate').classList.remove('visible');
            document.getElementById('infoAffiliate').innerText='- Rp 0';
            document.getElementById('pctAffiliate').innerText='0%';
        }
    }

    document.getElementById('infoTotalAdmin').innerText=formatRp(totalAdminInkAff);
    document.getElementById('pctTotalAdmin').innerText=`(${(ns>0?(totalAdminInkAff/ns)*100:0).toFixed(1)}%)`;
    const denom=1-aPct-(acosP/100)-affPct-(oprP/100)-(tnpm/100);

    const floor=denom>0&&tHpp>0?Math.ceil(tHpp/denom):0;
    document.getElementById('hargaFloor').innerText=floor>0?formatRp(floor):'—';

    // ── BANNER REKOMENDASI ────────────────────────────────────────────────────
    const banner=document.getElementById('alertBanner');
    const emptyPanel=document.getElementById('rightPanelEmpty');
    if(ns===0){banner.style.display='none';if(emptyPanel)emptyPanel.style.display='block';return;}
    if(emptyPanel)emptyPanel.style.display='none';

    // Kalkulasi angka spesifik
    const maxHppSatuan=denom>0?(ns*denom)/qty:0;
    const currentMarginBase=ns>0?1-(tHpp/ns)-aPct-affPct:0;
    const maxAcos=(currentMarginBase-(oprP/100)-(tnpm/100))*100;
    const maxOpr=(currentMarginBase-(acosP/100)-(tnpm/100))*100;
    const selisihHarga=floor>0?(floor-ns):0;

    // Jalur Margin: naikkan harga ke floor
    // Gunakan rasio (cairBersih/ns) sebagai proxy net-cair rate
    const cairRate = ns>0 ? cairBersih/ns : 0;
    const labaJalurMargin = floor>0 ? (floor*cairRate) - tHpp - (targetOrder>0?Math.floor(oprI/targetOrder):Math.floor(floor*(oprP/100))) - Math.floor(floor*(acosP/100)) : 0;
    const kapasitasOrder=maxOrderBln>0?maxOrderBln:targetOrder;
    const labaBlnMargin=labaJalurMargin*kapasitasOrder;

    // Jalur QTY: harga tetap, berapa order untuk capai target laba
    const targetLabaTotal=parseVal('targetLabaRp')||0;
    const orderUntukTargetLaba=laba>0&&targetLabaTotal>0?Math.ceil(targetLabaTotal/laba):0;
    const labaBlnQty=laba*kapasitasOrder;

    const txtAcos=maxAcos>0
        ?`Pangkas ACOS ke <b>${maxAcos.toFixed(1)}%</b> (naikkan ROAS ke <b>${(100/maxAcos).toFixed(1)}</b>)`
        :`<span style="color:#991b1b">Matikan iklan 100% pun tetap rugi!</span>`;
    const txtOpr=maxOpr>0
        ?`Efisiensi OPR ke <b>${maxOpr.toFixed(1)}%</b>`
        :`<span style="color:#991b1b">Potong OPR 100% pun tetap rugi!</span>`;

    // ── KONDISI AMAN ──────────────────────────────────────────────────────────
    if(laba>=0 && npm>=tnpm){
        banner.className='alert-banner alert-ok';
        banner.style.display='flex';

        const ruangTurunHarga = floor>0 ? ns-floor : 0;
        // roasMin: ROAS minimum agar laba >= target NPM
        // cairBersih - hpp - opr - (ns/roasMin) - fAff >= ns*tnpm/100
        // ns/roasMin <= cairBersih - hpp - opr - fAff - ns*tnpm/100
        // roasMin >= ns / (cairBersih - hpp - opr - fAff - ns*tnpm/100)
        const sisaUntukAds = cairBersih - tHpp - opr - (d.existsAff?0:fAff) - ns*(tnpm/100);
        const roasMin = sisaUntukAds > 0 ? Math.ceil(ns / sisaUntukAds) : 0;

        // Simulasi ROAS naik → laba naik
        const roasOpt1 = roas + 1;
        const acosOpt1 = 100 / roasOpt1;
        const adsOpt1 = ns * (acosOpt1/100);
        const labaOpt1 = cairBersih - tHpp - opr - adsOpt1 - (d.existsAff?0:Math.floor(ns*(pAff/100)));
        const tambahanLabaRoas = labaOpt1 - laba;
        const tambahanLabaBulanRoas = tambahanLabaRoas * kapasitasOrder;

        // Untuk 2x laba bulanan = kapasitasOrder * 2 (volume naik 2x)
        const orderUntuk2xLaba = kapasitasOrder * 2;

        // Simulasi HPP turun 5%
        const hppOpt = tHpp * 0.95;
        const labaHppOpt = cairBersih - hppOpt - opr - ads - (d.existsAff?0:Math.floor(ns*(pAff/100)));
        const tambahanLabaHpp = labaHppOpt - laba;

        // Status margin: aman tipis atau aman tebal?
        const marginStatus = npm < 8
            ? `<span style="color:#92400e;background:#fef9c3;padding:2px 8px;border-radius:4px;font-size:0.8em;font-weight:700;">⚠️ TIPIS — Rentan guncangan biaya</span>`
            : `<span style="color:#166534;background:#dcfce7;padding:2px 8px;border-radius:4px;font-size:0.8em;font-weight:700;">💪 SEHAT — Margin aman</span>`;

        // Peringatan jika margin tipis
        const warningTipis = npm < 8 ? `
            <div style="background:#fffbeb;border:1px solid #fcd34d;border-radius:8px;padding:10px;margin-bottom:10px;font-size:0.8em;color:#92400e;line-height:1.6;">
                <b>⚡ Perhatian:</b> NPM ${npm.toFixed(1)}% sangat tipis. Kenaikan biaya sekecil apapun (HPP, admin, voucher) bisa langsung membalik ke rugi. 
                Prioritaskan optimasi ROAS sebelum scale up volume.
            </div>` : '';

        banner.innerHTML=`
            <div style="width:100%;display:flex;flex-direction:column;flex:1;min-height:0;">
                <div style="display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:12px;flex-wrap:wrap;gap:6px;flex-shrink:0;">
                    <div style="font-size:1.15em;font-weight:800;">✅ TARGET TERCAPAI!</div>
                    ${marginStatus}
                </div>

                ${warningTipis}

                <div style="display:flex;flex-direction:column;gap:10px;flex:1;min-height:0;margin-bottom:12px;">
                    <!-- REKOMENDASI 1: Optimasi ROAS -->
                    <div style="background:#fff;border-radius:12px;padding:16px;border-left:5px solid #3730a3;box-shadow:0 1px 4px rgba(0,0,0,0.06);flex:1;min-height:0;display:flex;flex-direction:column;justify-content:center;">
                        <div style="font-size:0.8em;font-weight:800;color:#3730a3;text-transform:uppercase;margin-bottom:8px;letter-spacing:0.5px;">🎯 Rekomendasi 1 — Optimasi ROAS</div>
                        <div style="font-size:0.95em;line-height:2;color:#222;">
                            ROAS sekarang: <b>${roas.toFixed(1)}</b> → target <b>${roasOpt1.toFixed(1)}</b><br>
                            ACOS turun: <b>${acosP.toFixed(1)}%</b> → <b>${acosOpt1.toFixed(1)}%</b><br>
                            Laba/order naik: <b>+${formatRp(tambahanLabaRoas)}</b><br>
                            Laba bulanan: <b>+${formatRp(tambahanLabaBulanRoas)}</b> ekstra<br>
                            <span style="color:#3730a3;font-weight:700;">→ Pangkas kata kunci iklan yg boros</span>
                        </div>
                    </div>

                    <!-- REKOMENDASI 2: Scale Volume -->
                    <div style="background:#fff;border-radius:12px;padding:16px;border-left:5px solid #0d9488;box-shadow:0 1px 4px rgba(0,0,0,0.06);flex:1;min-height:0;display:flex;flex-direction:column;justify-content:center;">
                        <div style="font-size:0.8em;font-weight:800;color:#0d9488;text-transform:uppercase;margin-bottom:8px;letter-spacing:0.5px;">📦 Rekomendasi 2 — Scale Volume</div>
                        <div style="font-size:0.95em;line-height:2;color:#222;">
                            Order sekarang: <b>${kapasitasOrder} order/bln</b><br>
                            Laba bulanan: <b>${formatRp(labaBlnQty)}</b><br>
                            Untuk 2× laba → butuh <b>${orderUntuk2xLaba} order/bln</b><br>
                            ${ruangTurunHarga>0?`Ruang turun harga: <b>${formatRp(ruangTurunHarga)}</b>`:'Harga sudah di floor — jangan turun lagi'}<br>
                            <span style="color:#0d9488;font-weight:700;">→ Tambah variasi / bundling produk</span>
                        </div>
                    </div>

                    <!-- REKOMENDASI 3: Efisiensi HPP -->
                    <div style="background:#fff;border-radius:12px;padding:16px;border-left:5px solid #ea580c;box-shadow:0 1px 4px rgba(0,0,0,0.06);flex:1;min-height:0;display:flex;flex-direction:column;justify-content:center;">
                        <div style="font-size:0.8em;font-weight:800;color:#ea580c;text-transform:uppercase;margin-bottom:8px;letter-spacing:0.5px;">💡 Rekomendasi 3 — Tekan HPP</div>
                        <div style="font-size:0.95em;line-height:2;color:#222;">
                            HPP sekarang: <b>${formatRp(tHpp)}</b>/order<br>
                            Jika HPP turun 5%: <b>${formatRp(hppOpt)}</b><br>
                            Laba/order naik: <b>+${formatRp(tambahanLabaHpp)}</b><br>
                            Laba bulanan naik: <b>+${formatRp(tambahanLabaHpp*kapasitasOrder)}</b><br>
                            <span style="color:#ea580c;font-weight:700;">→ Negosiasi supplier / beli lebih banyak</span>
                        </div>
                    </div>
                </div>

                <!-- RINGKASAN KONDISI TOKO -->
                <div style="background:#f0fdf4;border:1px solid #86efac;border-radius:12px;padding:14px;display:grid;grid-template-columns:repeat(2,1fr);gap:10px;text-align:center;flex-shrink:0;">
                    <div>
                        <div style="font-size:0.72em;color:#166534;font-weight:700;text-transform:uppercase;margin-bottom:4px;">Laba/Order</div>
                        <div style="font-weight:800;color:#14532d;font-size:1.1em;">${formatRp(laba)}</div>
                    </div>
                    <div>
                        <div style="font-size:0.72em;color:#166534;font-weight:700;text-transform:uppercase;margin-bottom:4px;">Laba Bulanan</div>
                        <div style="font-weight:800;color:#14532d;font-size:1.1em;">${formatRp(labaBlnQty)}</div>
                    </div>
                    <div>
                        <div style="font-size:0.72em;color:#166534;font-weight:700;text-transform:uppercase;margin-bottom:4px;">ROAS Min Aman</div>
                        <div style="font-weight:800;color:#14532d;font-size:1.1em;">${roasMin>0?roasMin:'—'}</div>
                    </div>
                    <div>
                        <div style="font-size:0.72em;color:#166534;font-weight:700;text-transform:uppercase;margin-bottom:4px;">Floor Price</div>
                        <div style="font-weight:800;color:#14532d;font-size:1.1em;">${floor>0?formatRp(floor):'—'}</div>
                    </div>
                </div>
            </div>`;
        updateUGPanel(floor, roas, acosP, ns, tHpp, opr, oprI, targetOrder, laba, kapasitasOrder);
        saveHistory(laba,npm,gpm,ns,qty);
        return;
    }

    // ── KONDISI BONCOS / TIPIS ────────────────────────────────────────────────
    const isBoncos=laba<0;
    banner.className=isBoncos?'alert-banner alert-danger':'alert-banner alert-warning';
    banner.style.display='flex';

    // Gap analisis untuk prioritas rekomendasi
    const adminPctReal = ns>0?(totalAdminInkAff/ns)*100:0;
    const gapPctBoncos = gpm - adminPctReal; // ruang setelah HPP & Admin
    const selisihHargaRp = floor>0?(floor-ns):0;

    // Simulasi naik harga ke floor
    const labaJalurHarga = floor>0 ? (floor*cairBersih/ns) - tHpp - (targetOrder>0?Math.floor(oprI/targetOrder):Math.floor(floor*(oprP/100))) - Math.floor(floor*(acosP/100)) - (d.existsAff?d.aff:Math.floor(floor*(pAff/100))) : 0;

    // Simulasi HPP turun 5%
    const hppTurun5 = tHpp*0.95;
    const labaHppTurun = cairBersih - hppTurun5 - opr - ads - (d.existsAff?d.aff:fAff);
    const hppMaksHematan = hpp - maxHppSatuan;

    // Simulasi ROAS naik
    const roasOpt = roas + 1;
    const acosOpt = 100/roasOpt;
    const labaRoasNaik = cairBersih - tHpp - opr - Math.floor(ns*(acosOpt/100)) - (d.existsAff?d.aff:fAff);

    // Prioritas: gap negatif → harga dulu; gap tipis → ROAS dulu
    const prioritasHarga = gapPctBoncos < 10;

    const headerText=isBoncos
        ?`🔴 RUGI <b>${formatRp(Math.abs(laba))}</b>/order — Gap GPM-Admin hanya <b>${gapPctBoncos.toFixed(1)}%</b>. Tindakan segera!`
        :`⚠️ MARGIN TIPIS <b>${npm.toFixed(1)}%</b> vs target <b>${tnpm}%</b> — Gap <b>${gapPctBoncos.toFixed(1)}%</b> mepet.`;

    banner.innerHTML=`
        <div style="width:100%;display:flex;flex-direction:column;flex:1;min-height:0;">
            <div style="font-size:1em;font-weight:800;margin-bottom:10px;flex-shrink:0;line-height:1.4;">${headerText}</div>

            <div style="display:flex;flex-direction:column;gap:8px;flex:1;min-height:0;">

                <!-- REC 1: NAIKAN HARGA (prioritas utama kalau gap mepet/negatif) -->
                <div style="background:rgba(255,255,255,0.7);border-radius:10px;padding:13px;border-left:5px solid #3730a3;flex:1;min-height:0;display:flex;flex-direction:column;justify-content:center;">
                    <div style="font-size:0.75em;font-weight:800;color:#3730a3;margin-bottom:6px;text-transform:uppercase;">
                        🎯 ${prioritasHarga?'⭐ PRIORITAS UTAMA — ':''}Naikkan Harga
                    </div>
                    <div style="font-size:0.88em;line-height:1.8;">
                        Harga sekarang: <b>${formatRp(ns)}</b><br>
                        Harga minimum aman: <b>${floor>0?formatRp(floor):'—'}</b>${selisihHargaRp>0?` <span style="color:#3730a3">(+${formatRp(selisihHargaRp)})</span>`:''}<br>
                        Laba setelah naik: <b>${formatRp(labaJalurHarga)}</b>/order<br>
                        <span style="color:#3730a3;font-weight:700;">→ Cara termudah & tercepat untuk balik untung</span>
                    </div>
                </div>

                <!-- REC 2: EFISIENSI ROAS -->
                <div style="background:rgba(255,255,255,0.7);border-radius:10px;padding:13px;border-left:5px solid #0d9488;flex:1;min-height:0;display:flex;flex-direction:column;justify-content:center;">
                    <div style="font-size:0.75em;font-weight:800;color:#0d9488;margin-bottom:6px;text-transform:uppercase;">
                        📊 ${!prioritasHarga?'⭐ PRIORITAS UTAMA — ':''}Efisiensi Iklan (ROAS)
                    </div>
                    <div style="font-size:0.88em;line-height:1.8;">
                        ROAS sekarang: <b>${roas.toFixed(1)}</b> (ACOS ${acosP.toFixed(1)}%)<br>
                        Target ROAS: <b>${roasOpt.toFixed(1)}</b> → ACOS jadi <b>${acosOpt.toFixed(1)}%</b><br>
                        Dampak laba: <b>${labaRoasNaik>=0?'+':''} ${formatRp(labaRoasNaik-laba)}</b>/order<br>
                        ${maxAcos>0
                            ?`ACOS maks aman: <b>${maxAcos.toFixed(1)}%</b> (ROAS min <b>${(100/maxAcos).toFixed(1)}</b>)`
                            :`<span style="color:#991b1b">Matikan iklan pun masih rugi — harga harus naik dulu</span>`
                        }<br>
                        <span style="color:#0d9488;font-weight:700;">→ Pangkas kata kunci boros, naikkan bid produk terlaris</span>
                    </div>
                </div>

                <!-- REC 3: TEKAN HPP -->
                <div style="background:rgba(255,255,255,0.7);border-radius:10px;padding:13px;border-left:5px solid #ea580c;flex:1;min-height:0;display:flex;flex-direction:column;justify-content:center;">
                    <div style="font-size:0.75em;font-weight:800;color:#ea580c;margin-bottom:6px;text-transform:uppercase;">💡 Tekan HPP Produk</div>
                    <div style="font-size:0.88em;line-height:1.8;">
                        HPP sekarang: <b>${formatRp(tHpp)}</b>/order<br>
                        ${hppMaksHematan > 0
                            ? `HPP maks aman: <b>${formatRp(maxHppSatuan)}</b> → perlu hemat <b>${formatRp(hppMaksHematan)}/pcs</b>`
                            : `HPP sudah di atas batas aman — <b>negosiasi supplier wajib</b>`
                        }<br>
                        Jika HPP turun 5%: laba naik <b>+${formatRp(labaHppTurun-laba)}</b>/order<br>
                        <span style="color:#ea580c;font-weight:700;">→ Nego supplier / beli volume lebih besar / cari alternatif</span>
                    </div>
                </div>

                <!-- OPR INFO (tidak bisa ditawar) -->
                <div style="background:rgba(255,243,205,0.5);border-radius:8px;padding:10px 13px;border-left:3px solid #fbbf24;flex-shrink:0;">
                    <div style="font-size:0.72em;font-weight:800;color:#92400e;margin-bottom:3px;">🏭 OPR — TARGET TETAP (tidak bisa ditawar)</div>
                    <div style="font-size:0.82em;color:#92400e;line-height:1.6;">
                        Beban OPR: <b>${oprP}%</b> dari omset = <b>${formatRp(opr)}</b>/order &nbsp;|&nbsp;
                        Target omset: <b>${formatRp(targetOrder>0?targetOrder*ns:0)}</b>/bln untuk tutup OPR Rp ${formatRp(oprI)}/bln
                    </div>
                </div>
            </div>
        </div>`;

    updateUGPanel(floor, roas, acosP, ns, tHpp, opr, oprI, targetOrder, laba, kapasitasOrder);
    saveHistory(laba,npm,gpm,ns,qty);
}


// ═══════════════════════════════════════════════════════
// RIGHT PANEL — Toggle antara Rekomendasi & UG
// ═══════════════════════════════════════════════════════
let _ugData = null; // simpan data UG terakhir

function switchRightPanel(view) {
    const vRek = document.getElementById('viewRekomen');
    const vUG  = document.getElementById('viewUG');
    const title = document.getElementById('rightPanelTitle');
    const btnR = document.getElementById('rpBtnRekomen');
    const btnU = document.getElementById('rpBtnUG');
    if (view === 'ug') {
        if(vRek) vRek.style.display = 'none';
        if(vUG)  vUG.style.display  = 'flex'; vUG.style.flexDirection = 'column';
        if(title) title.innerText = '⚡ UG — Upgrade Margin';
        if(btnR){ btnR.style.background='#fff'; btnR.style.color='#ee4d2d'; }
        if(btnU){ btnU.style.background='#3730a3'; btnU.style.color='#fff'; }
    } else {
        if(vRek) vRek.style.display = 'flex'; vRek.style.flexDirection = 'column';
        if(vUG)  vUG.style.display  = 'none';
        if(title) title.innerText = '💡 Rekomendasi Aksi';
        if(btnR){ btnR.style.background='#ee4d2d'; btnR.style.color='#fff'; }
        if(btnU){ btnU.style.background='#fff'; btnU.style.color='#3730a3'; }
    }
}

function toggleRightPanelUG() {
    const vUG = document.getElementById('viewUG');
    const isUGVisible = vUG && vUG.style.display !== 'none';
    switchRightPanel(isUGVisible ? 'rekomen' : 'ug');
}

function updateUGPanel(floor, roas, acosP, ns, tHpp, opr, oprI, targetOrder, laba, kapasitasOrder) {
    _ugData = { floor, roas, acosP, ns, tHpp, opr, oprI, targetOrder, laba, kapasitasOrder };

    // Tampilkan tab buttons di bawah panel
    const rpTabs = document.getElementById('rpTabBtns');
    if (rpTabs && ns > 0) rpTabs.style.display = 'block';

    // Jalur 1: Naikan Harga
    const ugHarga = document.getElementById('ugHargaTarget');
    const ugSelisih = document.getElementById('ugSelisihHarga');
    if (ugHarga) {
        if (floor > 0 && ns > 0) {
            const selisih = floor - ns;
            ugHarga.innerHTML = '<b>Floor Price:</b> ' + formatRp(floor);
            if (ugSelisih) ugSelisih.innerText = selisih > 0
                ? '↑ Perlu naik ' + formatRp(selisih) + ' dari harga sekarang'
                : '✓ Harga sudah di atas floor price';
        } else {
            ugHarga.innerText = 'Proses pesanan dulu';
            if (ugSelisih) ugSelisih.innerText = '';
        }
    }

    // Jalur 2: Naikkan ROAS
    const ugRoas = document.getElementById('ugRoasTarget');
    const ugRoasImp = document.getElementById('ugRoasImpact');
    if (ugRoas && ns > 0) {
        const roasTarget = roas + 1;
        const acosTarget = (100 / roasTarget).toFixed(1);
        const savedAds = ns * (acosP/100) - ns * (acosTarget/100);
        ugRoas.innerHTML = 'ROAS <b>' + roas.toFixed(1) + '</b> → <b>' + roasTarget.toFixed(1) + '</b> (ACOS ' + acosTarget + '%)';
        if (ugRoasImp) ugRoasImp.innerText = '↑ Hemat iklan +' + formatRp(savedAds) + '/order';
    } else if (ugRoas) {
        ugRoas.innerText = 'Proses pesanan dulu';
        if (ugRoasImp) ugRoasImp.innerText = '';
    }

    // Jalur 3: OPR Target Omset
    const ugOpr = document.getElementById('ugOprTarget');
    const ugOprImp = document.getElementById('ugOprImpact');
    if (ugOpr && ns > 0) {
        const targetOmset = targetOrder * ns;
        ugOpr.innerHTML = 'Target Omset: <b>' + formatRp(targetOmset) + '</b>/bln';
        if (ugOprImp) ugOprImp.innerText = targetOrder + ' order/bln → OPR ' + formatRp(opr) + '/order';
    } else if (ugOpr) {
        ugOpr.innerText = 'Set OPR di HPP & Ads Setup';
        if (ugOprImp) ugOprImp.innerText = '';
    }
}

function saveHistory(laba,npm,gpm,ns,qty){
    if(d.cair>0){
        let hist=JSON.parse(localStorage.getItem('riwayat')||'[]');
        hist.unshift({
            waktu:new Date().toLocaleString('id-ID'),
            harga:formatRp(d.harga*qty),
            laba:formatRp(laba),
            npm:npm.toFixed(1),
            gpm:gpm.toFixed(1),
            labaRaw:laba,
            npmRaw:npm
        });
        if(hist.length>20) hist=hist.slice(0,20);
        localStorage.setItem('riwayat',JSON.stringify(hist));
    }
}

// SIMULATOR V7.7 — Input IDR/Persen, Harga Jual Otomatis
function resetInputSim(){
    document.getElementById('inputHppNaik').value = '0';
    document.getElementById('inputAdminNaik').value = '0';
    updateSimulator();
}

function parseRawSim(val){ return smartParseNumber(val); }

const parseRaw = v => smartParseNumber(v);

function updateSimulator(){
    const hppNaikRp = parseRawSim(document.getElementById('inputHppNaik').value);
    const admNaikPct = parseFloat(document.getElementById('inputAdminNaik').value)||0;

    const qty = d.qty||1;
    const hppBase = parseVal('hppSatuan') * qty;
    const acosP = smartParseNumber(document.getElementById('acos').value);
    const oprI = parseVal('oprTotalMonth');
    const oprP = parseVal('oprPctTarget');
    const ns = d.subtotal * qty;
    const ta = d.admin + d.layanan + d.proses;
    const aP = ns > 0 ? ta/ns : 0.138;

    if(ns === 0){document.getElementById('simAlert').style.display='block'; document.getElementById('outputHargaBaru').value='— Proses pesanan dulu —'; return;}
    document.getElementById('simAlert').style.display='none';

    const calc = (hppExtra, admExtra, nsOverride) => {
        const hppN = hppBase + hppExtra;
        const nsN = nsOverride !== null ? nsOverride : ns;
        const admN = nsN * (aP + admExtra);
        const adsN = nsN > 0 ? nsN*(acosP/100) : 0;
        const tgOrd = nsN > 0 ? Math.ceil((oprP > 0 ? oprI/(oprP/100) : 0)/nsN) : 0;
        const oprN = tgOrd > 0 ? Math.floor(oprI/tgOrd) : Math.floor(nsN*(oprP/100));
        const rasioPotonganLain = ns > 0 ? (d.aff + d.kampanye)/ns : 0;
        const cairN = nsN - admN - (nsN * rasioPotonganLain);
        const lN = cairN - hppN - oprN - adsN;
        return {hpp:hppN, ns:nsN, adm:admN, ads:adsN, laba:lN, npm:nsN>0?(lN/nsN)*100:0};
    };

    // Skenario 1: kondisi sekarang (tidak ada perubahan)
    const s1 = calc(0, 0, null);

    // Skenario 2: HPP naik (Rp) + Admin naik (%), harga tetap
    const s2 = calc(hppNaikRp, admNaikPct/100, null);

    // Skenario 3: Hitung harga jual baru agar NPM sama dengan Skenario 1
    // Rumus: nsNew * (1 - aP - admNaikPct/100 - acosP/100 - rasio_lain - oprP/100) = hppBase + hppNaikRp + oprNominal
    // Karena oprNominal tergantung nsNew, kita gunakan persentase: oprP/100 * nsNew
    const rasioPotonganLain = ns > 0 ? (d.aff + d.kampanye)/ns : 0;
    const npmTarget = s1.ns > 0 ? s1.laba/s1.ns : 0; // NPM Skenario 1
    const denomS3 = 1 - (aP + admNaikPct/100) - (acosP/100) - rasioPotonganLain - (oprP/100) - npmTarget;
    let nsNew = denomS3 > 0 ? Math.ceil((hppBase + hppNaikRp) / denomS3) : 0;
    // Jangan turun dari NS saat ini
    if(nsNew < ns) nsNew = ns;
    const s3 = calc(hppNaikRp, admNaikPct/100, nsNew);

    // Tampilkan harga jual baru di input readonly
    const hargaBaru = nsNew > 0 ? nsNew : 0;
    document.getElementById('outputHargaBaru').value = hargaBaru > 0 ? hargaBaru.toLocaleString('id-ID') : '—';

    const fill = (pfx, s) => {
        document.getElementById(pfx+'harga').innerText = formatRp(s.ns);
        document.getElementById(pfx+'hpp').innerText = formatRp(s.hpp);
        document.getElementById(pfx+'admin').innerText = formatRp(s.adm);
        document.getElementById(pfx+'ads').innerText = formatRp(s.ads);
        document.getElementById(pfx+'laba').innerText = formatRp(s.laba);
        const el = document.getElementById(pfx+'npm');
        el.innerText = 'NPM: ' + s.npm.toFixed(1) + '%';
        el.className = 'sim-npm ' + (s.npm>=8?'npm-ok':s.npm>=3?'npm-warn':'npm-bad');
    };
    fill('sc1', s1); fill('sc2', s2); fill('sc3', s3);
}

// REVERSE PRICING V7.5
function hitungReverse(){
    const hpp=parseRaw(document.getElementById('revHpp').value);
    const tLaba=parseRaw(document.getElementById('revTargetLaba').value);
    const tNpm=smartParseNumber(document.getElementById('revNpm').value)||0;
    const admP=(smartParseNumber(document.getElementById('revAdminPct').value)||24)/100;
    const roas=smartParseNumber(document.getElementById('revRoas').value)||4;
    const affP=(smartParseNumber(document.getElementById('revAff').value)||0)/100;
    const adsPct=roas>0?1/roas:0.25;
    const oprI=parseVal('oprTotalMonth');
    const oprP=parseVal('oprPctTarget');
    const dP_opr = oprP / 100; 
    
    const dA = 1 - admP - adsPct - affP - dP_opr;
    const dN = 1 - admP - adsPct - affP - dP_opr - (tNpm/100);
    
    let hDariLaba=0, hDariNpm=0;
    if(dA>0&&(tLaba>0||hpp>0)) hDariLaba=Math.ceil((tLaba+hpp)/dA);
    if(dN>0&&hpp>0) hDariNpm=Math.ceil(hpp/dN);
    
    const hRek=Math.max(hDariLaba,hDariNpm);
    
    if(hRek<=0){document.getElementById('revHargaJual').innerText='Rp —';return;}
    
    const oprNominal = hRek * dP_opr;
    const lR = hRek * (1 - admP - adsPct - affP) - hpp - oprNominal;
    const npmR=hRek>0?(lR/hRek)*100:0, gpmR=hRek>0?((hRek-hpp)/hRek)*100:0;
    const tOrd=hRek>0&&oprP>0&&oprI>0?Math.ceil((oprI/(oprP/100))/hRek):0;
    
    document.getElementById('revHargaJual').innerText=formatRp(hRek);
    document.getElementById('revGpm').innerText=gpmR.toFixed(1)+'%';
    document.getElementById('revNpmResult').innerText=npmR.toFixed(1)+'%';
    document.getElementById('revLabaOrder').innerText=formatRp(lR);
    document.getElementById('revAdminRp').innerText=formatRp(hRek*admP);
    document.getElementById('revAdsRp').innerText=formatRp(hRek*adsPct);
    document.getElementById('revOmsetBln').innerText=tOrd>0?tOrd+' order':' —';
}

// RIWAYAT
function renderHistory(){
    const list=document.getElementById('historyList');
    const hist=JSON.parse(localStorage.getItem('riwayat')||'[]');
    if(!hist.length){list.innerHTML='<li style="text-align:center;color:#999;font-size:0.82em;padding:24px;">Belum ada riwayat.</li>';return;}
    list.innerHTML=hist.map(h=>{
        const cls=h.npmRaw>=8?'h-ok':h.npmRaw>=0?'h-warn':'h-bad';
        const lbl=h.npmRaw>=8?'Aman':h.npmRaw>=0?'Tipis':'Rugi';
        return `<li class="history-item"><div><div style="font-weight:700;">${h.harga}</div><div style="color:#888;font-size:0.82em;">${h.waktu}</div></div><div style="text-align:right;"><div>Laba: <b>${h.laba}</b></div><div style="margin-top:3px;"><span class="history-badge ${cls}">${lbl} &middot; NPM ${h.npm}%</span></div></div></li>`;
    }).join('');
}
function hapusRiwayat(){localStorage.removeItem('riwayat');renderHistory();}

// ═══════════════════════════════════════════════════════
// MASTER DATA TOKO
// ═══════════════════════════════════════════════════════
let hppMaster = [];

function muatMasterData() {
    const saved = localStorage.getItem('masterDataToko');
    if (saved) { hppMaster = JSON.parse(saved); renderHppTable(); }
    const ops = JSON.parse(localStorage.getItem('masterOps') || '{}');
    if (ops.roas) document.getElementById('md_roas').value = ops.roas;
    if (ops.opr_total) document.getElementById('md_opr_total').value = ops.opr_total;
    if (ops.opr_pct) document.getElementById('md_opr_pct').value = ops.opr_pct;
    if (ops.npm) document.getElementById('md_npm').value = ops.npm;
    if (ops.max_order) document.getElementById('md_max_order').value = ops.max_order;
    if (ops.roas) updateAcosDisplay();
}

function updateAcosDisplay() {
    const roas = parseFloat(document.getElementById('md_roas').value) || 0;
    document.getElementById('md_acos').value = roas > 0 ? (100/roas).toFixed(1) + '%' : '—';
}

function syncMasterToMain() {
    updateAcosDisplay();
    const roas = parseFloat(document.getElementById('md_roas').value) || 0;
    const oprTotal = parseVal('md_opr_total');
    const oprPct = parseFloat(document.getElementById('md_opr_pct').value) || 0;
    const npm = parseFloat(document.getElementById('md_npm').value) || 0;
    const maxOrder = parseVal('md_max_order');
    if (roas > 0) { document.getElementById('roas').value = roas; syncAds('roas'); }
    if (oprTotal > 0) { document.getElementById('oprTotalMonth').value = oprTotal.toLocaleString('id-ID'); }
    if (oprPct > 0) { document.getElementById('oprPctTarget').value = oprPct; }
    if (npm > 0) { document.getElementById('targetNpm').value = npm; }
    if (maxOrder > 0) { document.getElementById('maxOrderBln').value = maxOrder.toLocaleString('id-ID'); }
    localStorage.setItem('masterOps', JSON.stringify({
        roas: document.getElementById('md_roas').value,
        opr_total: document.getElementById('md_opr_total').value,
        opr_pct: document.getElementById('md_opr_pct').value,
        npm: document.getElementById('md_npm').value,
        max_order: document.getElementById('md_max_order').value
    }));
    cleanAndHitung();
    updateRasioDashboard();
}

function toggleHppPaste() {
    const area = document.getElementById('hppPasteArea');
    area.style.display = area.style.display === 'none' ? 'block' : 'none';
    if (area.style.display === 'block') document.getElementById('hppPasteInput').focus();
}

function prosesHppPaste() {
    const raw = document.getElementById('hppPasteInput').value.trim();
    if (!raw) return;
    const lines = raw.split('\n').filter(l => l.trim());
    const newData = [];
    for (const line of lines) {
        const cols = line.split('\t');
        if (cols.length < 5) continue;
        const hpp = parseFloat(String(cols[4]).replace(/[Rp\.\s]/g,'').replace(',','.')) || 0;
        if (!cols[2] || hpp <= 0) continue;
        newData.push({ skuInduk:(cols[0]||'').trim(), namaProduk:(cols[1]||'').trim(), refSku:(cols[2]||'').trim(), namaVariasi:(cols[3]||'').trim(), hpp });
    }
    if (newData.length === 0) { alert('Data tidak terbaca. Pastikan copy langsung dari Google Sheet (5 kolom tab-separated).'); return; }
    newData.forEach(n => {
        const idx = hppMaster.findIndex(h => h.refSku === n.refSku);
        if (idx >= 0) hppMaster[idx] = n; else hppMaster.push(n);
    });
    localStorage.setItem('masterDataToko', JSON.stringify(hppMaster));
    document.getElementById('hppPasteInput').value = '';
    toggleHppPaste();
    renderHppTable();
}

let hppEditMode = false;

function renderHppTable() {
    const wrap = document.getElementById('hppTableWrap');
    const empty = document.getElementById('hppEmptyState');
    const search = document.getElementById('hppSearchWrap');
    const count = document.getElementById('hppCount');
    const footer = document.getElementById('hppTableFooter');
    const btnToggleEdit = document.getElementById('btnToggleEdit');
    const btnResetHpp = document.getElementById('btnResetHpp');
    if (hppMaster.length === 0) {
        wrap.style.display='none'; empty.style.display='block'; search.style.display='none';
        if(footer) footer.style.display='none';
        if(btnToggleEdit) btnToggleEdit.style.display='none';
        if(btnResetHpp) btnResetHpp.style.display='none';
        count.innerText='0 SKU'; return;
    }
    wrap.style.display='block'; empty.style.display='none'; search.style.display='block';
    if(footer) footer.style.display='block';
    if(btnToggleEdit) btnToggleEdit.style.display='inline-block';
    if(btnResetHpp) btnResetHpp.style.display='inline-block';
    count.innerText=hppMaster.length+' SKU';
    renderHppRows(hppMaster);
}

function toggleEditMode() {
    hppEditMode = !hppEditMode;
    const btn = document.getElementById('btnToggleEdit');
    const thAksi = document.getElementById('thAksi');
    if(btn) btn.innerHTML = hppEditMode ? '✅ Selesai' : '✏️ Edit';
    if(btn) btn.style.borderColor = hppEditMode ? '#16a34a' : '#3730a3';
    if(btn) btn.style.color = hppEditMode ? '#16a34a' : '#3730a3';
    if(thAksi) thAksi.style.display = hppEditMode ? 'table-cell' : 'none';
    renderHppRows(hppMaster);
}

function resetOperasional() {
    if (!confirm('Reset data ROAS, Operasional, dan Target Margin? Data HPP produk tidak terpengaruh.')) return;
    localStorage.removeItem('masterOps');
    ['md_roas','md_opr_total','md_opr_pct','md_npm','md_max_order'].forEach(id => { document.getElementById(id).value = ''; });
    document.getElementById('md_acos').value = '—';
}

function resetHppSaja() {
    if (!confirm('Hapus semua data HPP produk? Data ROAS & Operasional tidak terpengaruh.')) return;
    hppMaster = [];
    localStorage.removeItem('masterDataToko');
    hppEditMode = false;
    renderHppTable();
}

function renderHppRows(data) {
    document.getElementById('hppTableBody').innerHTML = data.map((h,i) => `
        <tr class="hpp-row">
            <td style="color:#bbb;font-size:0.78em;">${i+1}</td>
            <td><span style="font-weight:700;color:#555;">${h.skuInduk}</span></td>
            <td style="color:#444;max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="${h.namaProduk}">${h.namaProduk}</td>
            <td><span class="sku-chip">${h.refSku}</span></td>
            <td style="color:#777;">${h.namaVariasi}</td>
            <td style="text-align:right;"><span class="hpp-val">${h.hpp.toLocaleString('id-ID')}</span></td>
            <td style="display:${hppEditMode?'table-cell':'none'};text-align:center;">
                <button onclick="editHppRow(${hppMaster.indexOf(h)})" style="background:none;border:none;cursor:pointer;font-size:1em;padding:2px 6px;" title="Edit HPP">✏️</button>
            </td>
        </tr>`).join('');
    // Update footer counter
    const footer = document.getElementById('hppTableFooter');
    if (footer) footer.textContent = `Menampilkan ${data.length} dari ${hppMaster.length} SKU`;
}

function filterHppTable() {
    const q = document.getElementById('hppSearch').value.toLowerCase().trim();
    const filtered = q === '' ? hppMaster : hppMaster.filter(h =>
        h.refSku.toLowerCase().includes(q) ||
        h.namaProduk.toLowerCase().includes(q) ||
        h.skuInduk.toLowerCase().includes(q) ||
        h.namaVariasi.toLowerCase().includes(q)
    );
    renderHppRows(filtered);
}

function editHppRow(idx) {
    const h = hppMaster[idx];
    const newHpp = prompt(`Edit HPP untuk ${h.refSku} (${h.namaVariasi})\nHPP saat ini: Rp ${h.hpp.toLocaleString('id-ID')}\n\nMasukkan HPP baru:`, h.hpp);
    if (newHpp === null) return;
    const val = parseFloat(String(newHpp).replace(/\./g,'').replace(',','.'));
    if (isNaN(val) || val <= 0) { alert('Nilai tidak valid'); return; }
    hppMaster[idx].hpp = val;
    localStorage.setItem('masterDataToko', JSON.stringify(hppMaster));
    renderHppTable();
}

function editMasterData() { toggleHppPaste(); }

function resetMasterData() {
    if (!confirm('Reset semua data toko (ROAS, OPR, Target, dan HPP)?')) return;
    hppMaster = [];
    localStorage.removeItem('masterDataToko');
    localStorage.removeItem('masterOps');
    hppEditMode = false;
    renderHppTable();
    ['md_roas','md_opr_total','md_opr_pct','md_npm','md_max_order'].forEach(id => { document.getElementById(id).value = ''; });
    document.getElementById('md_acos').value = '—';
}

function lookupHppFromMaster(refSku) {
    if (!refSku || hppMaster.length === 0) return null;
    const found = hppMaster.find(h => h.refSku.toLowerCase() === refSku.toLowerCase().trim());
    return found ? found.hpp : null;
}


// ══════════════════════ RASIO KEUANGAN ══════════════════════
let rkData = {
    income: null,      // {totalPendapatan, totalPenghasilan, adminBreakdown:{ams,admin,layanan,proses,kampanye}}
    order1: null,      // {rows:[{noPesanan, refSku, qty, hppMatch}], totalOrder, totalHpp, totalPenghasilan}
    order2: null,      // same — bulan lalu
    ads: null,         // {totalAds} — real cost iklan (total top-up saldo iklan bulan ini)
    performa: null,    // {produkArr, totalKunjungan, totalOrderDibuat, cvrStore, produkAktif}
    periode: '',
    _incomeNoPesanan: new Set(),
    _incomeRawRows: [],
    _orderRawMap: {}
};

// ── XLSX Parser (pakai SheetJS via CDN) ──────────────────────
function loadSheetJS(cb) {
    if (window.XLSX) { cb(); return; }
    const s = document.createElement('script');
    s.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
    s.onload = cb;
    document.head.appendChild(s);
}

function readXlsx(file, cb) {
    loadSheetJS(() => {
        const reader = new FileReader();
        reader.onload = e => {
            const wb = XLSX.read(e.target.result, {type:'array'});
            cb(wb);
        };
        reader.readAsArrayBuffer(file);
    });
}

function readCsv(file, cb) {
    const reader = new FileReader();
    reader.onload = e => cb(e.target.result);
    reader.readAsText(file);
}

// ── HANDLE UPLOAD ─────────────────────────────────────────────
function handleRasioUpload(type, input) {
    const file = input.files[0];
    if (!file) return;
    const box = document.getElementById('box' + type.charAt(0).toUpperCase() + type.slice(1));

    if (type === 'ads') {
        readCsv(file, text => {
            const clean = text.replace(/^\uFEFF/, '');
            const lines = clean.split('\n');

            // ════════════════════════════════════════════════════
            // LAPIS 2 — CATEGORY-BASED PARSING (tahan banting)
            // Tidak bergantung pada nama deskripsi persis.
            // Kategorikan by SIFAT transaksi, bukan nama string.
            // ════════════════════════════════════════════════════
            const CAT = {
                // Dipotong dari penghasilan Shopee — sudah include PPN
                TOPUP_PENGHASILAN: ['dari penghasilan'],
                // Top-up manual dari kantong seller — belum include PPN → × 1.11
                TOPUP_MANUAL: ['isi saldo otomatis', 'top up', 'topup', 'isi saldo', 'saldo iklan dan bonus'],
                // Gratis dari Shopee — BUKAN uang keluar → SKIP
                BONUS_SKIP: ['bonus saldo', 'bonus iklan', 'cashback', 'gratis saldo', 'reward', 'hadiah', 'komisi afiliasi'],
                // Pengeluaran saldo iklan — tracking saja, bukan kas keluar
                IKLAN_SPEND: ['iklan produk', 'product ad', 'ads spend', 'deduction for', 'campaign spend',
                              'flash sale fee', 'iklan toko', 'iklan pencarian', 'biaya iklan'],
            };

            function kategorikan(desc) {
                const d = desc.toLowerCase().trim();
                // Cek BONUS dulu (prioritas tinggi — jangan salah hitung sebagai topup)
                if (CAT.BONUS_SKIP.some(k => d.includes(k))) return 'BONUS';
                // Cek TOPUP dari Penghasilan
                if (CAT.TOPUP_PENGHASILAN.some(k => d.includes(k))) return 'TOPUP_PENGHASILAN';
                // Cek TOPUP Manual — pastikan bukan "dari penghasilan"
                if (CAT.TOPUP_MANUAL.some(k => d.includes(k))) return 'TOPUP_MANUAL';
                // Cek Pengeluaran Iklan
                if (CAT.IKLAN_SPEND.some(k => d.includes(k))) return 'IKLAN_SPEND';
                return 'UNKNOWN';
            }

            const PPN = 1.11;
            let adSpend = 0;
            let topupPenghasilan = 0, topupManual = 0, bonusShopee = 0, ppnAmount = 0, iklanTerpakai = 0;
            // LAPIS 3 — Warning: kumpulkan transaksi tidak dikenali
            const unknownTrx = [];

            // Cari header: fuzzy match baris yang ada 'deskripsi' DAN 'jumlah'
            let hdrIdx = -1;
            lines.forEach((l, i) => {
                const ll = l.toLowerCase();
                if (ll.includes('deskripsi') && ll.includes('jumlah')) hdrIdx = i;
            });

            if (hdrIdx >= 0) {
                // LAPIS 1 — Fuzzy column index
                const hdrs = lines[hdrIdx].split(',').map(h => h.trim().toLowerCase());
                const fuzzyIdx = (...keys) => {
                    for (const k of keys) {
                        const i = hdrs.findIndex(h => h.includes(k));
                        if (i >= 0) return i;
                    }
                    return -1;
                };
                const iDesc = fuzzyIdx('deskripsi', 'description', 'keterangan', 'nama transaksi');
                const iJml  = fuzzyIdx('jumlah', 'amount', 'nominal', 'nilai');
                const iCat  = fuzzyIdx('catatan', 'note', 'remark', 'detail');

                for (let i = hdrIdx + 1; i < lines.length; i++) {
                    const cols = lines[i].split(',');
                    if (cols.length < 2) continue;
                    const desc = iDesc >= 0 ? (cols[iDesc] || '').trim() : '';
                    if (!desc) continue;
                    const raw = iJml >= 0 ? (cols[iJml] || '').replace(/[^0-9.\-]/g, '') : '';
                    const val = parseFloat(raw);
                    if (isNaN(val) || val === 0) continue;

                    const kategori = kategorikan(desc);

                    switch (kategori) {
                        case 'IKLAN_SPEND':
                            if (val < 0) iklanTerpakai += Math.abs(val);
                            break;

                        case 'TOPUP_PENGHASILAN':
                            if (val > 0) {
                                // Angka di file sudah include PPN — tidak dikali lagi
                                topupPenghasilan += val;
                                adSpend += val;
                            }
                            break;

                        case 'TOPUP_MANUAL':
                            if (val > 0) {
                                // Kas riil keluar = val × PPN 11%
                                const kasKeluar = val * PPN;
                                topupManual += val;
                                ppnAmount += kasKeluar - val;
                                adSpend += kasKeluar;
                            }
                            break;

                        case 'BONUS':
                            // Gratis dari Shopee — catat tapi TIDAK masuk adSpend
                            if (val > 0) bonusShopee += val;
                            break;

                        case 'UNKNOWN':
                            // LAPIS 3 — kumpulkan untuk warning
                            unknownTrx.push({ desc, val });
                            break;
                    }
                }
            } else {
                // Fallback: tidak ada header → coba parse naif
                lines.forEach(l => {
                    const lower = l.toLowerCase();
                    const cols  = l.split(',');
                    const vals  = cols.map(c => parseFloat(c.replace(/[^0-9.\-]/g,''))).filter(n => !isNaN(n));
                    if (CAT.TOPUP_PENGHASILAN.some(k => lower.includes(k)) && vals.some(v => v > 0))
                        adSpend += Math.max(...vals.filter(v => v > 0));
                    else if (CAT.TOPUP_MANUAL.some(k => lower.includes(k)) && vals.some(v => v > 0)
                             && !CAT.BONUS_SKIP.some(k => lower.includes(k)))
                        adSpend += Math.max(...vals.filter(v => v > 0)) * PPN;
                });
            }

            // ── Koreksi: gunakan isiSaldo dari Income jika sudah diupload ─────────
            // Income lebih akurat (cut-off bulanan) vs CSV (cut-off harian)
            // SELALU gunakan isiSaldo Income bila tersedia — tanpa threshold
            const isiSaldoIncome = rkData.income?.isiSaldo || 0;
            if (isiSaldoIncome > 0) {
                adSpend = adSpend - topupPenghasilan + isiSaldoIncome;
                topupPenghasilan = isiSaldoIncome;
            }

            const totalAds = Math.round(adSpend);
            rkData.ads = {
                totalAds, rawAds: adSpend, iklanTerpakai: Math.round(iklanTerpakai),
                topupPenghasilan: Math.round(topupPenghasilan),
                topupManual: Math.round(topupManual),
                bonusShopee: Math.round(bonusShopee),
                ppnAmount: Math.round(ppnAmount),
                unknownTrx,
            };

            // ── LAPIS 3 — Render warning transaksi tidak dikenali ─────────────────
            const warnAdsEl = document.getElementById('warnUnknownAds');
            const warnContainer = document.getElementById('rk_st_ads')?.parentElement;
            if (unknownTrx.length > 0) {
                const totalUnknown = unknownTrx.reduce((s, t) => s + Math.abs(t.val), 0);
                let warnEl = warnAdsEl;
                if (!warnEl && warnContainer) {
                    warnEl = document.createElement('div');
                    warnEl.id = 'warnUnknownAds';
                    warnEl.style.cssText = 'margin:8px 0;padding:10px 14px;background:#fffbeb;border:1.5px solid #f59e0b;border-radius:8px;font-size:0.78em;color:#92400e;';
                    warnContainer.appendChild(warnEl);
                }
                if (warnEl) {
                    warnEl.innerHTML = `⚠️ <b>${unknownTrx.length} jenis transaksi CSV belum dikenali</b> (total: ${formatRp(totalUnknown)})<br>
                    <span style="color:#b45309;">Diasumsikan <b>tidak masuk</b> komponen IKLAN. Periksa manual:<br>
                    ${unknownTrx.map(t => `• "${t.desc}" = ${formatRp(Math.abs(t.val))}`).join('<br>')}
                    </span>`;
                }
            } else {
                const warnEl = document.getElementById('warnUnknownAds');
                if (warnEl) warnEl.remove();
            }

            const sisaSaldo = Math.round(adSpend - iklanTerpakai);
            const ppnInfo = ppnAmount > 0 ? ` | PPN: ${formatRp(Math.round(ppnAmount))}` : '';
            setBoxUploaded(document.getElementById('boxAds'), `✓ ${formatRp(totalAds)}`);
            document.getElementById('statusAds').innerText = `✓ ${formatRp(totalAds)} real cost`;
            document.getElementById('rk_st_ads').innerText =
                `Top-up: ${formatRp(totalAds)}${ppnInfo} | Terpakai: ${formatRp(Math.round(iklanTerpakai))} | Sisa saldo: ${formatRp(Math.max(0,sisaSaldo))}`;
            document.getElementById('rk_st_ads').style.color = '#166534';
            updateRasioDashboard();
        });
        return;
    }

            // Cari baris header: Urutan,Waktu,Deskripsi,Jumlah,...
            let hdrIdx = -1;
            lines.forEach((l, i) => {
                if (l.toLowerCase().includes('deskripsi') && l.toLowerCase().includes('jumlah')) hdrIdx = i;
            });

            // ── KONSEP PERHITUNGAN ADS (UPDATED — PPN-AWARE) ────────────────────────
            // Real cost iklan = kas riil yang keluar dari kantong seller untuk iklan,
            // sudah termasuk PPN 11% Shopee Ads (berlaku sejak 2022).
            readXlsx(file, wb => {
        if (type === 'income') parseIncomeSheet(wb, box);
        else if (type === 'order1') parseOrderSheet(wb, 1, box);
        else if (type === 'order2') parseOrderSheet(wb, 2, box);
        else if (type === 'performa') parsePerformaSheet(wb, box);
        updateRasioDashboard();
    });
}

// ── TAMPILKAN NAMA FILE DI BOX ────────────────────────────────
function showFileName(inputEl, boxId) {
    const file = inputEl.files[0];
    if (!file) return;
    const box = document.getElementById(boxId);
    if (!box) return;
    let nameEl = box.querySelector('.rk-upload-filename');
    if (!nameEl) {
        nameEl = document.createElement('div');
        nameEl.className = 'rk-upload-filename';
        nameEl.style.cssText = 'font-size:0.6em;color:#3730a3;font-weight:700;margin-top:3px;word-break:break-all;text-align:center;padding:0 4px;line-height:1.4;';
        box.appendChild(nameEl);
    }
    // Tampilkan nama file — potong jika terlalu panjang
    const name = file.name;
    nameEl.textContent = name.length > 32 ? name.substring(0,30) + '…' : name;
    nameEl.title = name;
}

function setBoxUploaded(box, msg) {
    if (!box) return;
    box.classList.add('uploaded');
    box.querySelector('.rk-upload-status').innerText = msg;
}

// ── PARSE PERFORMA PRODUK (parentskudetail) ───────────────────
function parsePerformaSheet(wb, box) {
    // File: Shopee Seller Center → Analisis Produk → Performa SKU Induk
    // Sheet name biasanya "Produk dengan Performa Terbaik" atau sheet pertama
    const shName = wb.SheetNames.find(n =>
        n.toLowerCase().includes('performa') ||
        n.toLowerCase().includes('produk') ||
        n.toLowerCase().includes('sku')
    ) || wb.SheetNames[0];

    const sh = wb.Sheets[shName];
    const rows = XLSX.utils.sheet_to_json(sh, { header: 1, defval: '' });
    if (!rows || rows.length < 2) return;

    // Cari baris header — row yang mengandung "Kunjungan" atau "Konversi"
    let hdrIdx = -1;
    rows.forEach((r, i) => {
        const joined = r.map(c => String(c)).join('|').toLowerCase();
        if (joined.includes('kunjungan') || joined.includes('konversi')) hdrIdx = i;
    });
    if (hdrIdx < 0) hdrIdx = 0;

    const headers = rows[hdrIdx].map(c => String(c).trim().toLowerCase());

    // Map kolom kunci dari file Shopee parentskudetail
    const colIdx = (keywords) => {
        for (const kw of keywords) {
            const i = headers.findIndex(h => h.includes(kw));
            if (i >= 0) return i;
        }
        return -1;
    };

    const iSku        = colIdx(['sku induk']);
    const iProduk     = colIdx(['produk']);
    const iKunjungan  = colIdx(['pengunjung produk (kunjungan)']);
    const iDiklik     = colIdx(['produk diklik']);
    const iCtrPct     = colIdx(['persentase klik']);
    const iCvrDibuat  = colIdx(['tingkat konversi pesanan (pesanan dibuat)']);
    const iCvrSiap    = colIdx(['tingkat konversi pesanan (pesanan siap dikirim)']);
    // PENTING: gunakan exact match — "pesanan dibuat" saja akan salah ke kolom [8] Total Penjualan
    const iOrderDibuat= colIdx(['pesanan dibuat']);  // akan di-override pakai exactCol
    const iOrderSiap  = colIdx(['pesanan siap dikirim']);
    const iPenjualan  = colIdx(['total penjualan']);
    const iDimasukkan = colIdx(['dimasukkan ke keranjang (produk)']);
    const iAtcRate    = colIdx(['tingkat konversi produk dimasukkan ke keranjang']);
    const iRepeat     = colIdx(['tingkat pesanan berulang (pesanan dibuat)']);

    // Cari kolom EXACT: "Pesanan Dibuat" dan "Pesanan Siap Dikirim" (bukan yang ada kata "Tingkat/Total/Produk" sebelumnya)
    const exactCol = (exactName) => headers.findIndex(h => h.trim() === exactName.toLowerCase().trim());
    const iOrderDibuatExact = exactCol('pesanan dibuat');
    const iOrderSiapExact   = exactCol('pesanan siap dikirim');
    const iKunjunganExact   = exactCol('pengunjung produk (kunjungan)');

    // Parse helper: angka format Indonesia (titik = ribuan, koma = desimal)
    const parseNum = v => {
        if (v === '' || v === '-' || v === undefined || v === null) return 0;
        if (typeof v === 'number') return v;
        // Hapus titik ribuan dulu, lalu ganti koma desimal → titik
        const s = String(v).replace(/%/g,'').replace(/\./g,'').replace(/,/g,'.').replace(/[^0-9\.\-]/g,'');
        return parseFloat(s) || 0;
    };
    const parsePct = v => {
        if (typeof v === 'number') return v;
        // Format Indonesia: "1,65%" → 1.65
        const s = String(v).replace('%','').replace(/\./g,'').replace(',','.').trim();
        return parseFloat(s) || 0;
    };

    // Kumpulkan data per SKU Induk (baris tanpa variasi = baris aggregate)
    const produkMap = {};
    for (let i = hdrIdx + 1; i < rows.length; i++) {
        const r = rows[i];
        // Hanya baris SKU Induk (kolom variasi kosong = '-' atau '')
        // Deteksi: Kode Variasi (col index ~3) kosong atau '-'
        const isAggregate = (String(r[3] || '').trim() === '-' || String(r[3] || '').trim() === '');
        if (!isAggregate) continue;

        const skuInduk = iSku >= 0 ? String(r[iSku] || '').trim() : '';
        if (!skuInduk || skuInduk === '') continue;

        const namaProduk = iProduk >= 0 ? String(r[iProduk] || '').substring(0, 40) : skuInduk;
        // Gunakan exactCol untuk kolom yang ambigu
        const _iKunjungan = iKunjunganExact >= 0 ? iKunjunganExact : iKunjungan;
        const _iOrder     = iOrderDibuatExact >= 0 ? iOrderDibuatExact : iOrderDibuat;
        const _iOrderSiap = iOrderSiapExact   >= 0 ? iOrderSiapExact   : iOrderSiap;

        const kunjungan  = _iKunjungan >= 0 ? parseNum(r[_iKunjungan]) : 0;
        const diklik     = iDiklik     >= 0 ? parseNum(r[iDiklik])     : 0;
        const ctrPct     = iCtrPct     >= 0 ? parsePct(r[iCtrPct])     : 0;
        const cvrDibuat  = iCvrDibuat  >= 0 ? parsePct(r[iCvrDibuat])  : 0;
        const cvrSiap    = iCvrSiap    >= 0 ? parsePct(r[iCvrSiap])    : 0;
        const orderDibuat= _iOrder     >= 0 ? parseNum(r[_iOrder])     : 0;
        const orderSiap  = _iOrderSiap >= 0 ? parseNum(r[_iOrderSiap]) : 0;
        const penjualan  = iPenjualan  >= 0 ? parseNum(r[iPenjualan])  : 0;
        const dimasukkan = iDimasukkan >= 0 ? parseNum(r[iDimasukkan]) : 0;
        const atcRate    = iAtcRate    >= 0 ? parsePct(r[iAtcRate])    : 0;
        const repeatPct  = iRepeat     >= 0 ? parsePct(r[iRepeat])     : 0;

        produkMap[skuInduk] = {
            skuInduk, namaProduk, kunjungan, diklik, ctrPct,
            cvrDibuat, cvrSiap, orderDibuat, orderSiap,
            penjualan, dimasukkan, atcRate, repeatPct
        };
    }

    const produkArr = Object.values(produkMap);
    const totalKunjungan  = produkArr.reduce((s, p) => s + p.kunjungan, 0);
    const totalOrderDibuat= produkArr.reduce((s, p) => s + p.orderDibuat, 0);
    const totalPenjualan  = produkArr.reduce((s, p) => s + p.penjualan, 0);
    const cvrStore = totalKunjungan > 0 ? (totalOrderDibuat / totalKunjungan) * 100 : 0;
    const produkAktif = produkArr.filter(p => p.orderDibuat > 0).length;

    // Hitung pembatalan dari selisih Dibuat vs Siap Kirim
    const totalOrderSiap      = produkArr.reduce((s,p)=>s+(p.orderSiap||0), 0);
    const totalBatalPerforma  = Math.max(0, totalOrderDibuat - totalOrderSiap);
    const pctBatalPerforma    = totalOrderDibuat > 0 ? (totalBatalPerforma / totalOrderDibuat) * 100 : 0;
    // Estimasi revenue hilang = cancel × AOV
    const aovPerforma         = totalOrderDibuat > 0 ? Math.round(totalPenjualan / totalOrderDibuat) : 0;
    const revenueHilang       = totalBatalPerforma * aovPerforma;

    // Simpan ke rkData
    rkData.performa = {
        produkArr, totalKunjungan, totalOrderDibuat, totalOrderSiap,
        totalPenjualan, cvrStore, produkAktif,
        totalBatalPerforma, pctBatalPerforma, aovPerforma, revenueHilang
    };
    // Sync ke ongkirBatal agar renderOngkirPembatalan bisa pakai
    if (!rkData.ongkirBatal) rkData.ongkirBatal = {};
    rkData.ongkirBatal.totalBatalPerforma  = totalBatalPerforma;
    rkData.ongkirBatal.pctBatalPerforma    = pctBatalPerforma;
    rkData.ongkirBatal.aovPerforma         = aovPerforma;
    rkData.ongkirBatal.revenueHilang       = revenueHilang;
    rkData.ongkirBatal.totalOrderDibuat    = totalOrderDibuat;
    rkData.ongkirBatal.totalOrderSiap      = totalOrderSiap;

    // Update UI
    setBoxUploaded(box || document.getElementById('boxPerforma'),
        `✓ ${produkAktif} produk aktif`);
    document.getElementById('statusPerforma').innerText =
        `✓ ${produkAktif} produk aktif`;

    const stEl = document.getElementById('rk_st_performa_detail');
    if (stEl) {
        stEl.innerText = `${produkArr.length} produk | ${totalKunjungan.toLocaleString('id-ID')} kunjungan | CVR toko ${cvrStore.toFixed(2)}%`;
        stEl.style.color = '#d97706';
    }
}

// ── PARSE INCOME SHEET ────────────────────────────────────────
function parseIncomeSheet(wb, box) {
    // File income Shopee: bisa punya sheet "Income" (per transaksi) DAN "Summary" (ringkasan)
    // Prioritas BENAR: Income per-transaksi DULU, baru fallback ke Summary
    const shIncome   = wb.SheetNames.find(n => n.toLowerCase() === 'income')  || null;
    const shSummary  = wb.SheetNames.find(n => n.toLowerCase() === 'summary') || null;
    const shFallback = wb.SheetNames[0];

    // Income (per transaksi) → Summary (ringkasan) → sheet pertama
    const shName = shIncome || shSummary || shFallback;
    const sh = wb.Sheets[shName];
    const rows = XLSX.utils.sheet_to_json(sh, {header:1, defval:''});

    // ── Deteksi format ─────────────────────────────────────────
    // Jika sheet "Income" ditemukan → pasti per-transaksi, skip cek Ringkasan
    // Jika tidak ada sheet Income → cek konten untuk deteksi Ringkasan
    const top20 = rows.slice(0, 20).map(r => r.map(c=>String(c)).join('|').toLowerCase());
    const isRingkasan = (!shIncome) && top20.some(t =>
        t.includes('ringkasan penghasilan') ||
        t.includes('1. total pendapatan') ||
        t.includes('laporan penghasilan')
    );

    let totalPendapatan=0, totalPenghasilan=0;
    let ams=0, admin=0, layanan=0, proses=0, kampanye=0;

    if (isRingkasan) {
        // ── FORMAT RINGKASAN (file Summary Shopee) ──────────────
        // Struktur: label di kolom 0 atau 1, nilai di kolom terakhir yang ada angka
        // Contoh: ["1. Total Pendapatan", "", "", 8245634]
        //         ["", "Biaya Komisi AMS", -314428, ""]
        let isiSaldo=0; // Biaya Isi Saldo Otomatis — bukan biaya riil, harus di-add back ke totalPenghasilan
        rows.forEach(r => {
            // Ambil semua teks non-kosong di row ini
            const lbl = r.map(c=>String(c).trim()).filter(c=>c).join(' ').toLowerCase();
            // Cari nilai numerik: ambil yang paling kanan / paling masuk akal
            let val = 0;
            for (let c = r.length-1; c >= 0; c--) {
                const n = parseFloat(r[c]);
                if (!isNaN(n) && n !== 0) { val = n; break; }
            }
            if (val === 0) return;

            if (lbl.includes('1. total pendapatan') || (lbl.includes('total pendapatan') && !lbl.includes('pengeluaran')))
                totalPendapatan = Math.abs(val);
            else if (lbl.includes('3. total yang dilepas') || lbl.includes('total yang dilepas') || lbl.includes('total dilepas'))
                totalPenghasilan = Math.abs(val);
            else if (lbl.includes('biaya komisi ams') || lbl.includes('komisi ams'))
                ams = Math.abs(val);
            else if (lbl.includes('biaya administrasi') && !lbl.includes('program'))
                admin = Math.abs(val);
            else if (lbl.includes('biaya layanan'))
                layanan = Math.abs(val);
            else if (lbl.includes('biaya proses pesanan') || lbl.includes('biaya proses'))
                proses = Math.abs(val);
            else if (lbl.includes('biaya kampanye'))
                kampanye = Math.abs(val);
            else if (lbl.includes('biaya isi saldo otomatis') || lbl.includes('isi saldo otomatis'))
                isiSaldo = Math.abs(val);
        });

        // Biaya Isi Saldo Otomatis: Shopee memotongnya dari dana dilepas sebagai transfer ke saldo iklan.
        // Ini BUKAN biaya operasional — harus ditambahkan kembali ke totalPenghasilan
        // agar angka Total Penghasilan = sama dengan hasil sum kolom "Total Penghasilan" per transaksi (File 5 xlsx).
        // TIDAK dimasukkan ke adminBreakdown / Rasio Admin.
        totalPenghasilan = totalPenghasilan + isiSaldo;

        rkData.income = { totalPendapatan, totalPenghasilan, adminBreakdown:{ams,admin,layanan,proses,kampanye}, isiSaldo };
        // Format ringkasan tidak punya No. Pesanan per baris — set kosong berarti no-filter di grafik
        rkData._incomeNoPesanan = new Set();

        if (totalPendapatan > 0)  document.getElementById('rk_totalPendapatan').value = Math.round(totalPendapatan).toLocaleString('id-ID');
        if (totalPenghasilan > 0) document.getElementById('rk_totalPenghasilan').value = Math.round(totalPenghasilan).toLocaleString('id-ID');
        if (admin+layanan+proses+kampanye > 0) document.getElementById('rk_adminTotal').value = Math.round(admin+layanan+proses+kampanye).toLocaleString('id-ID');
        if (ams > 0) document.getElementById('rk_amsTotal').value = Math.round(ams).toLocaleString('id-ID');

        setBoxUploaded(document.getElementById('boxIncome'), `✓ Ringkasan`);
        document.getElementById('statusIncome').innerText = `✓ Ringkasan`;
        document.getElementById('rk_st_income').innerText = `Ringkasan — ${formatRp(totalPenghasilan)} dilepas`;
        document.getElementById('rk_st_income').style.color = '#166534';
        // Tampilkan peringatan: format Ringkasan tidak didukung oleh Analisis Biaya Layanan
        const warnElR = document.getElementById('incomeFormatWarn');
        if (warnElR) warnElR.style.display = '';
        setTimeout(syncBiayaStatusBar, 0);

    } else {
        // ── FORMAT PER TRANSAKSI (Data Income per baris pesanan) ─
        let hdrIdx = -1;
        rows.forEach((r, i) => {
            const joined = r.map(c=>String(c)).join('|').toLowerCase();
            if (joined.includes('no. pesanan') && joined.includes('total penghasilan')) hdrIdx = i;
        });
        if (hdrIdx < 0) rows.forEach((r,i) => { if (hdrIdx<0 && r.filter(c=>c!=='').length>15) hdrIdx=i; });
        if (hdrIdx < 0) { console.warn('Format income tidak dikenali'); return; }

        const headers = rows[hdrIdx].map(h=>String(h).trim());
        // LAPIS 1 — Fuzzy column matching: cari kolom by daftar keyword, ambil yang pertama cocok
        const fuzzyCol = (...keys) => {
            for (const k of keys) {
                const i = headers.findIndex(h => h.toLowerCase().includes(k.toLowerCase()));
                if (i >= 0) return i;
            }
            return -1;
        };
        const iNoPesanan = fuzzyCol('No. Pesanan','Nomor Pesanan','Order ID','No Pesanan','No.Pesanan');
        const iTotal     = fuzzyCol('Total Penghasilan','Total Income','Net Income','Dana Diterima');
        const iHarga     = fuzzyCol('Harga Asli Produk','Harga Asli','Original Price','Harga Produk');
        const iAMS       = fuzzyCol('Biaya Komisi AMS','Komisi AMS','AMS Fee','AMS Commission','Shopee Ads Commission');
        const iAdmin     = fuzzyCol('Biaya Administrasi','Administration Fee','Admin Fee','Biaya Admin');
        const iLayanan   = fuzzyCol('Biaya Layanan','Service Fee','Layanan');
        const iProses    = fuzzyCol('Biaya Proses Pesanan','Order Processing Fee','Processing Fee','Biaya Proses');
        const iKampanye  = fuzzyCol('Biaya Kampanye','Campaign Fee','Kampanye');
        const iIsiSaldo  = fuzzyCol('Biaya Isi Saldo Otomatis','Isi Saldo Otomatis','Auto Top Up Fee','Top Up Fee');
        const iOngkirSeller     = fuzzyCol('Promo Gratis Ongkir dari Penjual','Gratis Ongkir Penjual','Seller Free Shipping');
        const iHematKirim       = fuzzyCol('Biaya Program Hemat Biaya Kirim','Hemat Biaya Kirim','Shipping Discount Program');
        const iOngkirSellerAlt  = fuzzyCol('Ongkos Kirim Ditanggung Penjual','Ongkir Ditanggung Penjual','Biaya Kirim Ditanggung Penjual');
        const iReturDana        = fuzzyCol('Pengembalian Dana ke Pembeli','Refund to Buyer','Retur Dana');
        const iReturOngkir      = fuzzyCol('Ongkos Kirim Pengembalian Barang','Return Shipping Fee','Retur Ongkir');
        const iNoPes = iNoPesanan >= 0 ? iNoPesanan : 1;
        let count=0, totalOngkirSeller=0, totalReturDana=0, totalReturOngkir=0, totalIsiSaldo=0;

        const noPesananSet = new Set();
        for (let i=hdrIdx+1; i<rows.length; i++) {
            const r=rows[i];
            const noPesanan=String(r[iNoPes]||'').trim();
            if (!noPesanan || noPesanan.toLowerCase().includes('pesanan')) continue;
            const n=idx=>idx>=0?Math.abs(parseFloat(r[idx])||0):0;
            const ha=n(iHarga), tp=n(iTotal);
            if (ha>0||tp>0) {
                totalPendapatan+=ha; totalPenghasilan+=tp;
                ams+=n(iAMS); admin+=n(iAdmin); layanan+=n(iLayanan); proses+=n(iProses); kampanye+=n(iKampanye);
                // Capture Isi Saldo per baris (akan di-add back setelah loop, TIDAK masuk adminBreakdown)
                if (iIsiSaldo >= 0) totalIsiSaldo += n(iIsiSaldo);
                // Ongkir seller: pakai kolom baru, fallback ke kolom lama
                if (iOngkirSeller >= 0) totalOngkirSeller += n(iOngkirSeller);
                else if (iOngkirSellerAlt >= 0) totalOngkirSeller += n(iOngkirSellerAlt);
                if (iHematKirim >= 0) totalOngkirSeller += n(iHematKirim);
                // Retur
                if (iReturDana >= 0) totalReturDana += n(iReturDana);
                if (iReturOngkir >= 0) totalReturOngkir += n(iReturOngkir);
                noPesananSet.add(noPesanan);
                count++;
            }
        }

        // Biaya Isi Saldo Otomatis: Shopee memotongnya dari kolom "Total Penghasilan" per transaksi.
        // Ini BUKAN biaya operasional — add-back agar totalPenghasilan = sama dengan File xlsx.
        // TIDAK dimasukkan ke adminBreakdown / Rasio Admin.
        totalPenghasilan = totalPenghasilan + totalIsiSaldo;

        // ── Simpan raw rows untuk Analisis Biaya Layanan ──
        const incomeRawRows = [];
        const BIAYA_NAMES_R = ['Biaya Komisi AMS','Biaya Administrasi','Biaya Layanan','Biaya Proses Pesanan','Biaya Kampanye','Biaya Program Hemat Biaya Kirim','Biaya Transaksi','Premi'];
        const colByExact = name => headers.findIndex(h => h.trim() === name);
        for (let ri=hdrIdx+1; ri<rows.length; ri++) {
            const r = rows[ri];
            const noPes = String(r[iNoPes]||'').trim();
            if (!noPes || noPes.toLowerCase().includes('pesanan')) continue;
            const nv = idx => idx>=0 ? (parseFloat(r[idx])||0) : 0;
            const omsetR = nv(iHarga), incomeR = nv(iTotal);
            if (omsetR===0 && incomeR===0) continue;
            const kodeVchrR = String(r[colByExact('Kode Voucher')]||'').trim();
            const decProg = (c, gofsVal) => { const s=String(c||'').trim(); if(s.startsWith('VIDEO-'))return'Video Extra'; if(s.startsWith('LIVE-'))return'Live Extra'; if(s.startsWith('SVC-'))return'Vchr Extra'; if(s.toUpperCase().includes('ALLE')&&s)return'Voucher Penjual'; if(!s && Math.abs(gofsVal||0)>0)return'GOFS Extra'; return'Tanpa Program'; };
            const totalDiskonR = nv(colByExact('Total Diskon Produk'));
            const diskonShopeeR = nv(colByExact('Diskon Produk dari Shopee'));
            const voucherSellerR = nv(colByExact('Voucher disponsor oleh Penjual'));
            const gofsR = nv(colByExact('Gratis Ongkir dari Shopee'));
            const biayaDetR = {}; let totalBiayaR = 0;
            BIAYA_NAMES_R.forEach(nm => { const v=nv(colByExact(nm)); biayaDetR[nm]=v; totalBiayaR+=v; });
            incomeRawRows.push({
                noPesanan:noPes, kodeVchr:kodeVchrR, program:decProg(kodeVchrR, gofsR),
                omset:omsetR, diskonSellerProduk:totalDiskonR-diskonShopeeR,
                voucherSeller:voucherSellerR, gofs:gofsR, income:incomeR,
                totalBiaya:totalBiayaR, biayaDetail:biayaDetR,
                totalPotonganSeller:(totalDiskonR-diskonShopeeR)+voucherSellerR,
            });
        }
        rkData._incomeRawRows = incomeRawRows;

        // ── Override totalPendapatan dari sheet Summary ─────────────────────────
        // Sheet Income per-transaksi menggunakan "Harga Asli Produk" (bruto sebelum diskon).
        // Nilai yang benar sebagai base rasio adalah "1. Total Pendapatan" dari sheet Summary
        // (= nett setelah diskon & voucher), sama dengan yang tertera di laporan ringkasan Shopee.
        if (shSummary) {
            try {
                const shSum = wb.Sheets[shSummary];
                const rowsSum = XLSX.utils.sheet_to_json(shSum, {header:1, defval:''});
                for (const r of rowsSum) {
                    const lbl = r.map(c=>String(c).trim()).filter(c=>c).join(' ').toLowerCase();
                    let val = 0;
                    for (let c = r.length-1; c >= 0; c--) {
                        const n = parseFloat(r[c]);
                        if (!isNaN(n) && n !== 0) { val = n; break; }
                    }
                    if (val === 0) continue;
                    if (lbl.includes('1. total pendapatan') || (lbl.includes('total pendapatan') && !lbl.includes('pengeluaran'))) {
                        totalPendapatan = Math.abs(val);
                        break;
                    }
                }
            } catch(e) { console.warn('Gagal baca Summary untuk totalPendapatan:', e); }
        }
        // ────────────────────────────────────────────────────────────────────────

        rkData.income = { totalPendapatan, totalPenghasilan, adminBreakdown:{ams,admin,layanan,proses,kampanye}, isiSaldo:totalIsiSaldo };
        rkData._incomeNoPesanan = noPesananSet;
        if (!rkData.ongkirBatal) rkData.ongkirBatal = {};
        rkData.ongkirBatal.totalOngkirSeller = totalOngkirSeller;
        rkData.ongkirBatal.totalReturDana    = totalReturDana;
        rkData.ongkirBatal.totalReturOngkir  = totalReturOngkir;
        if (totalPendapatan>0)  document.getElementById('rk_totalPendapatan').value=Math.round(totalPendapatan).toLocaleString('id-ID');
        if (totalPenghasilan>0) document.getElementById('rk_totalPenghasilan').value=Math.round(totalPenghasilan).toLocaleString('id-ID');
        if (admin+layanan+proses+kampanye>0) document.getElementById('rk_adminTotal').value=Math.round(admin+layanan+proses+kampanye).toLocaleString('id-ID');
        if (ams>0) document.getElementById('rk_amsTotal').value=Math.round(ams).toLocaleString('id-ID');

        setBoxUploaded(document.getElementById('boxIncome'), `✓ ${count} transaksi`);
        document.getElementById('statusIncome').innerText=`✓ ${count} transaksi`;
        document.getElementById('rk_st_income').innerText=`${count} transaksi — ${formatRp(totalPenghasilan)} cair`;
        document.getElementById('rk_st_income').style.color='#166534';
        // Sembunyikan warning Ringkasan (kalau sebelumnya pernah muncul)
        const warnElT = document.getElementById('incomeFormatWarn');
        if (warnElT) warnElT.style.display = 'none';
        setTimeout(syncBiayaStatusBar, 0);
    }

    // ── Recalculate iklan jika CSV Ads sudah diupload sebelum Income ──────────
    // isiSaldo dari Income SELALU lebih akurat — langsung ganti tanpa threshold
    if (rkData.ads && rkData.income?.isiSaldo > 0) {
        const isiSaldoIncome = rkData.income.isiSaldo;
        const oldTopup = rkData.ads.topupPenghasilan || 0;
        // Selalu recalculate — tidak ada threshold minimum selisih
        rkData.ads.totalAds = Math.round(rkData.ads.totalAds - oldTopup + isiSaldoIncome);
        rkData.ads.topupPenghasilan = isiSaldoIncome;
        // Update tampilan box Ads
        const boxAds = document.getElementById('boxAds');
        if (boxAds) setBoxUploaded(boxAds, `✓ ${formatRp(rkData.ads.totalAds)}`);
        const stAds = document.getElementById('statusAds');
        if (stAds) stAds.innerText = `✓ ${formatRp(rkData.ads.totalAds)} real cost`;
        const rkStAds = document.getElementById('rk_st_ads');
        if (rkStAds) rkStAds.innerText = `Top-up: ${formatRp(rkData.ads.totalAds)} | Terpakai: ${formatRp(rkData.ads.iklanTerpakai||0)}`;
    }

    updateRasioDashboard();
}

// ── PARSE ORDER SHEET ─────────────────────────────────────────
function parseOrderSheet(wb, num, box) {
    let sh = wb.Sheets['Data Order'] || wb.Sheets[wb.SheetNames.find(n => n.toLowerCase().includes('order')) || wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sh, {header:1, defval:''});

    let hdrIdx = 0;
    rows.forEach((r,i) => {
        const joined = String(r[0]||'').toLowerCase();
        if (joined.includes('no. pesanan') || joined.includes('nomor pesanan') ||
            joined.includes('order id')    || joined.includes('no pesanan')) hdrIdx=i;
    });
    const headers = rows[hdrIdx].map(h=>String(h).trim());

    // LAPIS 1 — Fuzzy column matching untuk Order file
    const fuzzyCol = (...keys) => {
        for (const k of keys) {
            const i = headers.findIndex(h => h.toLowerCase().includes(k.toLowerCase()));
            if (i >= 0) return i;
        }
        return -1;
    };
    const iNo       = Math.max(fuzzyCol('No. Pesanan','Nomor Pesanan','Order ID','No Pesanan'), 0);
    const iSku      = fuzzyCol('Nomor Referensi SKU','Ref SKU','SKU Ref','Reference SKU','SKU ID');
    const iQty      = fuzzyCol('Jumlah','Qty','Quantity','Kuantitas');
    const iRetQty   = fuzzyCol('Returned quantity','Qty Retur','Return Qty','Jumlah Retur');
    const iNamaProd = fuzzyCol('Nama Produk','Product Name','Nama Barang');
    const iVariasi  = fuzzyCol('Nama Variasi','Variation','Varian','Variasi');
    const iWaktu    = fuzzyCol('Waktu Pesanan Dibuat','Order Time','Waktu Order','Tanggal Pesanan');
    const iWaktuByr = fuzzyCol('Waktu Pembayaran Dilakukan','Payment Time','Waktu Bayar','Tanggal Bayar');
    const iKota     = fuzzyCol('Kota','City','Kabupaten');
    const iProvinsi = fuzzyCol('Provinsi','Province');
    const iStatus   = fuzzyCol('Status Pesanan','Order Status','Status');

    // Set No. Pesanan valid dari Income (untuk filter grafik)
    const incomeNoPesananSet = rkData._incomeNoPesanan || new Set();

    let totalOrder=0, totalHpp=0, matchCount=0, noMatchCount=0;
    let totalBatal=0, totalRetur=0;
    // seenOrder: untuk hitung jumlah pesanan unik (1 No. Pesanan = 1 order)
    // seenHpp: untuk hitung HPP per baris SKU (1 No. Pesanan + SKU = 1 baris HPP)
    // Ini mendukung: 1 pesanan bisa punya banyak SKU berbeda & qty > 1
    const seenOrder   = new Set();
    const seenHpp     = new Set(); // key: noPesanan + '||' + refSku
    const varianCount = {};
    const jamCount    = {};
    const kotaCount   = {};
    const tanggalCount = {}; // agregasi per tanggal untuk grafik

    for (let i=hdrIdx+1; i<rows.length; i++) {
        const r=rows[i];
        const noPesanan=String(r[iNo]||'').trim();
        if (!noPesanan) continue;

        const refSku = String(r[iSku]||'').trim();
        const retQty = parseInt(r[iRetQty])||0;
        const qty    = Math.max(0,(parseInt(r[iQty])||1)-retQty);
        if (retQty > 0) totalRetur += retQty;

        // ── Deteksi Pembatalan ──
        const statusPesanan = iStatus >= 0 ? String(r[iStatus]||'').toLowerCase() : '';
        if (statusPesanan.includes('batal') || statusPesanan.includes('cancel')) {
            if (!seenOrder.has(noPesanan)) { totalBatal++; seenOrder.add(noPesanan); }
            continue;
        }

        // ── Hitung Order unik (1 No. Pesanan = 1 order) ──
        if (!seenOrder.has(noPesanan)) {
            seenOrder.add(noPesanan);
            if(qty > 0) totalOrder++;

            // ── TANGGAL PEMBAYARAN (per order, bukan per baris) ──
            const waktuByrStr = iWaktuByr >= 0 ? String(r[iWaktuByr] || '') : '';
            const waktuFallback = iWaktu >= 0 ? String(r[iWaktu] || '') : '';
            const tglStr = (waktuByrStr || waktuFallback).split(' ')[0];
            if (tglStr && tglStr.includes('-')) {
                const valid = incomeNoPesananSet.size === 0 || incomeNoPesananSet.has(noPesanan);
                if (valid) tanggalCount[tglStr] = (tanggalCount[tglStr] || 0) + 1;
            }

            // ── JAM RAMAI (per order) ──
            if (iWaktu>=0) {
                const waktuStr = String(r[iWaktu]||'');
                const tPart = waktuStr.split(' ')[1] || waktuStr;
                const jamNum = parseInt(tPart.split(':')[0]);
                if (!isNaN(jamNum) && jamNum>=0 && jamNum<24) {
                    const label = String(jamNum).padStart(2,'0')+':00 – '+String(jamNum+1).padStart(2,'0')+':00';
                    jamCount[label] = (jamCount[label]||0) + 1;
                }
            }

            // ── KOTA / PROVINSI (per order) ──
            const kotaVal    = iKota>=0     ? String(r[iKota]||'').trim()     : '';
            const provinsiVal= iProvinsi>=0 ? String(r[iProvinsi]||'').trim() : '';
            const lokasi     = kotaVal || provinsiVal;
            if (lokasi) kotaCount[lokasi] = (kotaCount[lokasi]||0) + 1;
        }

        if(qty<=0) continue;

        // ── VARIAN TERLARIS — akumulasi qty per SKU (semua baris) ──
        const varianKey = refSku || (iNamaProd>=0 ? String(r[iNamaProd]||'').trim().substring(0,20) : '');
        if (varianKey) varianCount[varianKey] = (varianCount[varianKey]||0) + qty;

        // ── HPP — per baris (No. Pesanan + SKU), dukung multi-SKU & qty > 1 ──
        // Tiap kombinasi noPesanan+refSku dihitung sekali dengan qty aktual baris tsb.
        // Ini menangani: 1 pesanan berisi MAYRA (qty 2) + LAVERA (qty 1) = HPP keduanya terhitung.
        const hppKey = noPesanan + '||' + refSku;
        if (!seenHpp.has(hppKey)) {
            seenHpp.add(hppKey);
            const sudahCair = incomeNoPesananSet.size === 0 || incomeNoPesananSet.has(noPesanan);
            const found = hppMaster.find(h=>h.refSku.toLowerCase()===refSku.toLowerCase());
            if (found && sudahCair) { totalHpp += found.hpp * qty; matchCount++; }
            else if (!found && sudahCair && refSku) { noMatchCount++; }
        }
    }

    // ── Simpan raw voucher map untuk Analisis Biaya Layanan ──
    const iVchrShopeeO = headers.findIndex(h=>h.trim()==='Voucher Ditanggung Shopee');
    const iVchrSellerO = headers.findIndex(h=>h.trim()==='Voucher Ditanggung Penjual');
    for (let oi=hdrIdx+1; oi<rows.length; oi++) {
        const r = rows[oi];
        const noP = String(r[iNo]||'').trim();
        if (!noP) continue;
        const pNum = v => { let s=String(v||'').replace(/[^\d.,-]/g,''); s=s.replace(/\./g,'').replace(',','.'); const n=parseFloat(s); return isNaN(n)?0:n; };
        if (!rkData._orderRawMap[noP]) {
            rkData._orderRawMap[noP] = {
                vchrShopee: iVchrShopeeO>=0 ? pNum(r[iVchrShopeeO]) : 0,
                vchrSeller: iVchrSellerO>=0 ? pNum(r[iVchrSellerO]) : 0,
            };
        }
    }

    const key = num===1?'order1':'order2';
    // Simpan set No. Pesanan yang ada di Data Order (untuk cross-check dengan Income)
    rkData[key] = { totalOrder, totalHpp, matchCount, noMatchCount, varianCount, jamCount, kotaCount, tanggalCount, orderNoPesananSet: seenOrder };
    if (num===1) {
        if (!rkData.ongkirBatal) rkData.ongkirBatal = {};
        rkData.ongkirBatal.totalBatal = totalBatal;
        rkData.ongkirBatal.totalRetur = totalRetur;
    }

    // ── DETEKSI & KOMPENSASI: No. Pesanan di Income tapi tidak ada di Data Order ──
    // Kasus: order bulan lalu yang baru cair bulan ini — user hanya upload Data Order bulan ini
    // Solusi: hitung HPP dari Data Order yang sudah diupload, lalu tampilkan peringatan
    if (num===1) {
        const incomeSet = rkData._incomeNoPesanan || new Set();
        const missingOrders = [...incomeSet].filter(no => !seenOrder.has(no));
        if (missingOrders.length > 0) {
            // Tampilkan peringatan bahwa ada order di Income yang tidak ada di Data Order
            setTimeout(() => {
                const warnId = 'warnMissingOrder';
                let warnEl = document.getElementById(warnId);
                if (!warnEl) {
                    warnEl = document.createElement('div');
                    warnEl.id = warnId;
                    warnEl.style.cssText = 'margin:8px 0;padding:10px 14px;background:#fffbeb;border:1.5px solid #f59e0b;border-radius:8px;font-size:0.78em;color:#92400e;';
                    const hppBox = document.getElementById('rk_hppTotal');
                    if (hppBox) hppBox.parentElement.parentElement.insertAdjacentElement('afterend', warnEl);
                }
                warnEl.innerHTML = `⚠️ <b>${missingOrders.length} pesanan di Income tidak ditemukan di Data Order</b><br>
                <span style="color:#b45309;">Kemungkinan order bulan lalu yang baru cair bulan ini. HPP dari pesanan tersebut <b>tidak terhitung</b>.<br>
                Solusi: sertakan juga Data Order bulan lalu saat upload agar HPP akurat 100%.</span><br>
                <span style="font-size:0.9em;color:#78350f;margin-top:4px;display:block;">No. Pesanan yang terlewat: ${missingOrders.slice(0,5).join(', ')}${missingOrders.length>5?' ...':''}</span>`;
            }, 100);
        } else {
            // Hapus warning kalau sudah lengkap
            const warnEl = document.getElementById('warnMissingOrder');
            if (warnEl) warnEl.remove();
        }
    }

    // Update field HPP Total = order1 + order2 (bulan lalu) — selalu recalculate setiap kali ada data baru
    const combinedHpp = (rkData.order1?.totalHpp || 0) + (rkData.order2?.totalHpp || 0);
    if (combinedHpp > 0) { document.getElementById('rk_hppTotal').value = Math.round(combinedHpp).toLocaleString('id-ID'); }

    const statusEl = num===1?'statusOrder1':'statusOrder2';
    const stEl     = num===1?'rk_st_order1':'rk_st_order2';
    const boxId    = num===1?'boxOrder1':'boxOrder2';

    setBoxUploaded(document.getElementById(boxId), '✓ '+totalOrder+' order');
    document.getElementById(statusEl).innerText='✓ '+totalOrder+' order';
    document.getElementById(stEl).innerText=totalOrder+' order, '+matchCount+' HPP matched';
    document.getElementById(stEl).style.color='#166634';

    const hppEl=document.getElementById('rk_st_hpp');
    if(hppEl) {
        const totalMatch = (rkData.order1?.matchCount||0) + (rkData.order2?.matchCount||0);
        const totalNoMatch = (rkData.order1?.noMatchCount||0) + (rkData.order2?.noMatchCount||0);
        hppEl.innerText=totalMatch+' matched, '+totalNoMatch+' unmatched';
        hppEl.style.color=totalNoMatch>0?'#b45309':'#166634';
    }

    // Setelah order2 diparse, re-check apakah warning "pesanan tidak ditemukan" masih relevan
    if (num===2) {
        const incomeSet = rkData._incomeNoPesanan || new Set();
        const order1Set = rkData.order1?.orderNoPesananSet || new Set();
        const order2Set = seenOrder;
        const allOrderSet = new Set([...order1Set, ...order2Set]);
        const missingOrders = [...incomeSet].filter(no => !allOrderSet.has(no));
        const warnEl = document.getElementById('warnMissingOrder');
        if (missingOrders.length === 0 && warnEl) {
            warnEl.innerHTML = '✅ <b>Semua pesanan di Income ditemukan di Data Order</b> — HPP sudah akurat 100%.';
            warnEl.style.background = '#f0fdf4';
            warnEl.style.borderColor = '#86efac';
            warnEl.style.color = '#166534';
            setTimeout(() => { if (warnEl.parentNode) warnEl.remove(); }, 4000);
        } else if (missingOrders.length > 0 && warnEl) {
            warnEl.innerHTML = `⚠️ <b>${missingOrders.length} pesanan di Income masih belum ada di Data Order</b><br>
            <span style="color:#b45309;">HPP belum 100% akurat. Tambahkan Data Order yang berisi pesanan ini.<br>
            No. Pesanan: ${missingOrders.slice(0,5).join(', ')}${missingOrders.length>5?' ...':''}</span>`;
        }
    }

    if (num===1) {
        renderVarianTerlaris(varianCount, totalOrder);
        renderJamRamai(jamCount);
        renderKotaTerbanyak(kotaCount, totalOrder);
    }
    updateRasioDashboard();
}

// ── RENDER VARIAN TERLARIS ─────────────────────────────────────
function renderVarianTerlaris(varianCount, totalOrder) {
    const el = document.getElementById('varianList');
    if (!el) return;
    const sorted = Object.entries(varianCount).sort((a,b)=>b[1]-a[1]).slice(0,5);
    if (sorted.length===0) { el.innerHTML='<div style="text-align:center;padding:14px;color:#bbb;font-size:0.78em;">Tidak ada data varian</div>'; return; }
    const maxQty = sorted[0][1];
    const medals = ['🥇','🥈','🥉'];
    el.innerHTML = sorted.map(([nama,qty],i)=>{
        const pct  = totalOrder>0 ? ((qty/totalOrder)*100).toFixed(1) : 0;
        const barW = maxQty>0 ? Math.round((qty/maxQty)*100) : 0;
        const badge= i<3 ? medals[i] : '<span style="font-size:0.75em;color:#aaa;font-weight:700;">#'+(i+1)+'</span>';
        const short= nama.length>34 ? nama.substring(0,32)+'…' : nama;
        return '<div style="padding:7px 12px;border-bottom:1px solid #f5f5f5;">'
            +'<div style="display:flex;align-items:center;gap:6px;margin-bottom:3px;">'
            +'<span style="font-size:0.85em;">'+badge+'</span>'
            +'<span style="font-size:0.73em;font-weight:700;color:#222;flex:1;line-height:1.3;" title="'+nama+'">'+short+'</span>'
            +'<span style="font-size:0.72em;font-weight:800;color:#ee4d2d;white-space:nowrap;">'+qty+' pcs</span>'
            +'</div>'
            +'<div style="display:flex;align-items:center;gap:6px;">'
            +'<div style="flex:1;height:5px;background:#f0f0f0;border-radius:3px;overflow:hidden;">'
            +'<div style="width:'+barW+'%;height:100%;background:#ee4d2d;border-radius:3px;"></div></div>'
            +'<span style="font-size:0.68em;color:#aaa;font-weight:600;min-width:34px;text-align:right;">'+pct+'%</span>'
            +'</div></div>';
    }).join('');
}

// ── RENDER JAM RAMAI ──────────────────────────────────────────
function renderJamRamai(jamCount) {
    const el = document.getElementById('jamRamaiList');
    if (!el) return;
    const sorted = Object.entries(jamCount).sort((a,b)=>b[1]-a[1]).slice(0,5);
    if (sorted.length===0) { el.innerHTML='<div style="text-align:center;padding:14px;color:#bbb;font-size:0.78em;">Kolom Waktu Pesanan tidak ditemukan di file</div>'; return; }
    const maxCnt = sorted[0][1];
    const total  = Object.values(jamCount).reduce((a,b)=>a+b,0);
    const cols   = ['#0d9488','#0f766e','#14b8a6','#2dd4bf','#5eead4'];
    el.innerHTML = sorted.map(([jam,cnt],i)=>{
        const barW = maxCnt>0 ? Math.round((cnt/maxCnt)*100) : 0;
        const pct  = total>0 ? ((cnt/total)*100).toFixed(0) : 0;
        return '<div style="margin-bottom:9px;">'
            +'<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:3px;">'
            +'<span style="font-size:0.76em;font-weight:700;color:#222;">'+jam+'</span>'
            +'<span style="font-size:0.72em;font-weight:800;color:#0d9488;">'+cnt+' order <span style="color:#aaa;font-weight:500;">('+pct+'%)</span></span>'
            +'</div>'
            +'<div style="height:6px;background:#f0fdfa;border-radius:4px;overflow:hidden;">'
            +'<div style="width:'+barW+'%;height:100%;background:'+(cols[i]||cols[4])+';border-radius:4px;"></div></div></div>';
    }).join('');
}

// ── RENDER KOTA / PROVINSI TERBANYAK ─────────────────────────
function renderKotaTerbanyak(kotaCount, totalOrder) {
    const el = document.getElementById('kotaList');
    if (!el) return;
    const sorted = Object.entries(kotaCount).sort((a,b)=>b[1]-a[1]).slice(0,6);
    if (sorted.length===0) { el.innerHTML='<div style="text-align:center;padding:14px;color:#bbb;font-size:0.78em;">Kolom Kota/Provinsi tidak ditemukan di file pesanan</div>'; return; }
    const maxCnt = sorted[0][1];
    el.innerHTML = sorted.map(([kota,cnt])=>{
        const pct  = totalOrder>0 ? ((cnt/totalOrder)*100).toFixed(1) : 0;
        const barW = maxCnt>0 ? Math.round((cnt/maxCnt)*100) : 0;
        return '<div style="padding:7px 12px;border-bottom:1px solid #f5f5f5;">'
            +'<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:3px;">'
            +'<span style="font-size:0.75em;font-weight:700;color:#222;">📍 '+kota+'</span>'
            +'<span style="font-size:0.72em;font-weight:800;color:#7c3aed;">'+cnt+' order</span>'
            +'</div>'
            +'<div style="display:flex;align-items:center;gap:6px;">'
            +'<div style="flex:1;height:5px;background:#f5f3ff;border-radius:3px;overflow:hidden;">'
            +'<div style="width:'+barW+'%;height:100%;background:#7c3aed;border-radius:3px;"></div></div>'
            +'<span style="font-size:0.68em;color:#aaa;font-weight:600;min-width:34px;text-align:right;">'+pct+'%</span>'
            +'</div></div>';
    }).join('');
}

// ── ADS MANUAL ────────────────────────────────────────────────
function toggleAdsManual() {
    const w=document.getElementById('adsManualWrap');
    w.style.display=w.style.display==='none'?'block':'none';
}
function prosesAdsManual() {
    const v=smartParseNumber(document.getElementById('adsManualVal').value);
    if(v>0){
        // Input manual = real cost iklan bulan ini (sudah termasuk PPN jika ada)
        // Masukkan nilai total kas yang keluar dari kantong (sudah PPN-inclusive)
        rkData.ads={totalAds:Math.round(v)};
        document.getElementById('statusAds').innerText=`✓ Manual: ${formatRp(Math.round(v))}`;
        document.getElementById('rk_st_ads').innerText=`Real cost iklan: ${formatRp(Math.round(v))}`;
        document.getElementById('rk_st_ads').style.color='#166634';
        updateRasioDashboard();
    }
}


// ── GRAFIK ORDER HARIAN ────────────────────────────────────────
function renderGrafikHarian() {
    const canvas = document.getElementById('chartOrderHarian');
    const emptyEl = document.getElementById('chartEmptyState');
    const tooltip = document.getElementById('chartTooltip');
    if (!canvas) return;

    const d1 = rkData.order1?.tanggalCount || {};
    const d2 = rkData.order2?.tanggalCount || {};
    if (Object.keys(d1).length === 0 && Object.keys(d2).length === 0) {
        if (emptyEl) emptyEl.style.display = 'flex';
        return;
    }
    if (emptyEl) emptyEl.style.display = 'none';

    // Bangun array 1–31 hari dari data yang ada
    // Ambil semua tanggal unik, deteksi bulan dari data
    const allDates1 = Object.keys(d1).sort();
    const allDates2 = Object.keys(d2).sort();

    // Tentukan range hari (1-31) dari bulan masing-masing
    const getDay = dateStr => parseInt((dateStr || '').split('-')[2]) || 0;
    const days = Array.from({length: 31}, (_, i) => i + 1);

    // Map tanggal ke hari
    const mapToDay = (dateCount) => {
        const byDay = {};
        Object.entries(dateCount).forEach(([dt, cnt]) => {
            const day = getDay(dt);
            if (day > 0) byDay[day] = (byDay[day] || 0) + cnt;
        });
        return byDay;
    };

    const byDay1 = mapToDay(d1);
    const byDay2 = mapToDay(d2);
    const maxDay1 = allDates1.length > 0 ? getDay(allDates1[allDates1.length-1]) : 31;
    const maxDay2 = allDates2.length > 0 ? getDay(allDates2[allDates2.length-1]) : 31;
    const maxDay  = Math.max(maxDay1, maxDay2, 28);

    const vals1 = days.slice(0, maxDay).map(d => byDay1[d] || 0);
    const vals2 = days.slice(0, maxDay).map(d => byDay2[d] || 0);
    const labels = days.slice(0, maxDay);

    const ctx = canvas.getContext('2d');
    const container = canvas.parentElement;
    const W = container ? container.clientWidth : (canvas.offsetWidth || 900);
    const H = 180;
    canvas.width  = W * window.devicePixelRatio;
    canvas.height = H * window.devicePixelRatio;
    canvas.style.width  = W + 'px';
    canvas.style.height = H + 'px';
    ctx.scale(window.devicePixelRatio, window.devicePixelRatio);

    const PAD = { top: 16, right: 20, bottom: 32, left: 36 };
    const cW = W - PAD.left - PAD.right;
    const cH = H - PAD.top  - PAD.bottom;

    const allVals = [...vals1, ...vals2].filter(v => v > 0);
    const maxVal  = allVals.length > 0 ? Math.max(...allVals) : 5;
    const yMax    = maxVal + Math.ceil(maxVal * 0.2) || 5;
    const n       = labels.length;
    const xStep   = n > 1 ? cW / (n - 1) : cW;

    const xPos = i => PAD.left + i * xStep;
    const yPos = v => PAD.top + cH - (v / yMax) * cH;

    // Clear
    ctx.clearRect(0, 0, W, H);

    // Grid lines
    const gridLines = 4;
    ctx.strokeStyle = '#f0f0f0';
    ctx.lineWidth = 1;
    for (let g = 0; g <= gridLines; g++) {
        const y = PAD.top + (cH / gridLines) * g;
        ctx.beginPath(); ctx.moveTo(PAD.left, y); ctx.lineTo(PAD.left + cW, y); ctx.stroke();
        const lbl = Math.round(yMax - (yMax / gridLines) * g);
        ctx.fillStyle = '#bbb'; ctx.font = '500 9px Plus Jakarta Sans, sans-serif';
        ctx.textAlign = 'right'; ctx.fillText(lbl, PAD.left - 5, y + 3);
    }

    // Draw line helper
    function drawLine(vals, color, dashed) {
        if (vals.every(v => v === 0)) return;
        ctx.beginPath();
        ctx.strokeStyle = color;
        ctx.lineWidth = 2;
        ctx.lineJoin = 'round';
        ctx.lineCap  = 'round';
        if (dashed) ctx.setLineDash([5, 4]); else ctx.setLineDash([]);
        vals.forEach((v, i) => {
            const x = xPos(i), y = yPos(v);
            i === 0 ? ctx.moveTo(x, y) : ctx.lineTo(x, y);
        });
        ctx.stroke();
        ctx.setLineDash([]);
        // Dots
        vals.forEach((v, i) => {
            if (v === 0) return;
            ctx.beginPath();
            ctx.arc(xPos(i), yPos(v), 3, 0, Math.PI * 2);
            ctx.fillStyle = color; ctx.fill();
        });
    }

    drawLine(vals2, '#cbd5e1', true);  // bulan lalu — abu dashed
    drawLine(vals1, '#ee4d2d', false); // bulan ini — merah solid

    // X-axis labels (tampilkan tiap 5 hari)
    ctx.fillStyle = '#999'; ctx.font = '500 9px Plus Jakarta Sans, sans-serif'; ctx.textAlign = 'center';
    labels.forEach((d, i) => {
        if (d === 1 || d % 5 === 0 || d === maxDay) {
            ctx.fillText(d, xPos(i), H - PAD.bottom + 14);
        }
    });

    // Interaktif tooltip
    const hitTest = (mx, my) => {
        let closest = -1, minDist = 20;
        labels.forEach((_, i) => {
            const dx = Math.abs(mx - xPos(i));
            if (dx < minDist) { minDist = dx; closest = i; }
        });
        return closest;
    };

    canvas._chartData = { vals1, vals2, labels, xPos, yPos, PAD };

    canvas.onmousemove = (e) => {
        const rect = canvas.getBoundingClientRect();
        const mx = e.clientX - rect.left;
        const my = e.clientY - rect.top;
        const idx = hitTest(mx, my);
        if (idx >= 0 && tooltip) {
            const v1 = vals1[idx], v2 = vals2[idx];
            const day = labels[idx];
            tooltip.style.display = 'block';
            tooltip.style.left = (e.clientX + 12) + 'px';
            tooltip.style.top  = (e.clientY - 10) + 'px';
            tooltip.innerHTML =
                `<div style="margin-bottom:3px;color:#94a3b8;font-size:0.85em;">Tanggal ${day}</div>` +
                `<div style="color:#ee4d2d;">● Bulan Ini: <b>${v1} order</b></div>` +
                (v2 > 0 ? `<div style="color:#94a3b8;">● Bulan Lalu: <b>${v2} order</b></div>` : '');
        }
    };
    canvas.onmouseleave = () => { if (tooltip) tooltip.style.display = 'none'; };
}

// ── SYNC STATUS DETAIL (Upload & Data tab) ───────────────────
function syncStatusDetail() {
    const fields = ['income','order1','order2','ads','hpp'];
    fields.forEach(k => {
        const src = document.getElementById('rk_st_' + k);
        const dst = document.getElementById('rk_st_' + k + '_detail');
        if (src && dst) {
            dst.innerText = src.innerText || '—';
            dst.style.color = src.style.color || '#bbb';
        }
    });
    // Update shortcut text di Overview
    const shortcut = document.getElementById('rk_uploadShortcutStatus');
    if (shortcut) {
        const hasIncome = (document.getElementById('rk_st_income')?.innerText || '').length > 1;
        const hasOrder  = (document.getElementById('rk_st_order1')?.innerText || '').length > 1;
        if (hasIncome && hasOrder) {
            shortcut.innerText = '✅ Data sudah diproses — lihat hasil di bawah';
        } else if (hasIncome || hasOrder) {
            shortcut.innerText = '⚠️ Data sebagian — upload & proses semua file';
        } else {
            shortcut.innerText = 'Belum ada data — upload dulu';
        }
    }
}

// ── UPDATE DASHBOARD ──────────────────────────────────────────
function updateRasioDashboard() {
    // Reset canvas jika tidak ada data order
    if (!rkData.order1 && !rkData.order2) {
        const canvas = document.getElementById('chartOrderHarian');
        if (canvas) { const ctx = canvas.getContext('2d'); ctx.clearRect(0,0,canvas.width,canvas.height); }
        const emptyEl = document.getElementById('chartEmptyState');
        if (emptyEl) emptyEl.style.display = 'flex';
    }
    const periode=document.getElementById('rk_periode').value||'—';
    const thEl=document.getElementById('rk_thPeriode');
    if(thEl) thEl.innerText=periode;

    // Kumpulkan nilai — prioritas: input manual, fallback: parsed xlsx
    const get=id=>smartParseNumber(document.getElementById(id).value);
    // Gunakan input manual jika diisi, jika tidak fallback ke rkData
    const _manualPendapatan = get('rk_totalPendapatan');
    const _manualPenghasilan = get('rk_totalPenghasilan');
    const _manualHpp = get('rk_hppTotal');
    const _manualOpr = get('rk_oprTotal');
    const _manualAdmin = get('rk_adminTotal');
    const _manualAms = get('rk_amsTotal');

    const totalPendapatan = _manualPendapatan > 0 ? _manualPendapatan : (rkData.income?.totalPendapatan || 0);
    const totalPenghasilan = _manualPenghasilan > 0 ? _manualPenghasilan : (rkData.income?.totalPenghasilan || 0);
    // HPP = order1 + order2 (bulan lalu yg cair bulan ini), kecuali ada input manual
    const _hppFromData = (rkData.order1?.totalHpp || 0) + (rkData.order2?.totalHpp || 0);
    const hpp = _manualHpp > 0 ? _manualHpp : _hppFromData;
    const opr = _manualOpr > 0 ? _manualOpr : (smartParseNumber(localStorage.getItem('oprTotalMonth') || '0'));
    const _adminFromData = rkData.income ? (
        (rkData.income.adminBreakdown?.admin || 0) +
        (rkData.income.adminBreakdown?.layanan || 0) +
        (rkData.income.adminBreakdown?.proses || 0) +
        (rkData.income.adminBreakdown?.kampanye || 0)
    ) : 0;
    const adminTotal = _manualAdmin > 0 ? _manualAdmin : _adminFromData;
    const ams = _manualAms > 0 ? _manualAms : (rkData.income?.adminBreakdown?.ams || 0);
    const iklan=rkData.ads?.totalAds||0;
    const isiSaldo=rkData.income?.isiSaldo||0;
    const totalOrder=rkData.order1?.totalOrder||0;
    const maxOrderBln = smartParseNumber(document.getElementById('maxOrderBln')?.value || '0');

    // Derived
    const aov=totalOrder>0?Math.round(totalPendapatan/totalOrder):0;
    const basketSize=totalOrder>0?(rkData.order1?.totalOrder/totalOrder).toFixed(1):1;
    const roasAktual=iklan>0?(totalPendapatan/iklan).toFixed(2):0;
    const acosAktual=totalPendapatan>0?(iklan/totalPendapatan)*100:0;
    const adminRasio=totalPendapatan>0?-(adminTotal/totalPendapatan)*100:0;
    const layanan=rkData.income?.adminBreakdown?.layanan||0;
    const adminFee=rkData.income?.adminBreakdown?.admin||0;
    const proses=rkData.income?.adminBreakdown?.proses||0;
    const kampanye=rkData.income?.adminBreakdown?.kampanye||0;
    // ── Rasio komponen admin: DIBULATKAN ke integer sesuai spreadsheet ALLEY ──
    const pctRound = (val, base) => base > 0 ? Math.round((val / base) * 100) : 0;
    const amsRasioPct      = pctRound(ams,      totalPendapatan);
    const adminFeeRasioPct = pctRound(adminFee, totalPendapatan);
    const layananRasioPct  = pctRound(layanan,  totalPendapatan);
    const prosesRasioPct   = pctRound(proses,   totalPendapatan);
    const kampanyeRasioPct = pctRound(kampanye, totalPendapatan);
    const isiSaldoRasioPct = pctRound(isiSaldo, totalPendapatan);
    // SPREADSHEET ALLEY: TOTAL ADMIN = AMS+Adm+Layanan+Proses+Kampanye SAJA
    // IsiSaldo tampil TERPISAH — tidak masuk total admin header
    const totalAdminKomponen=ams+adminFee+layanan+proses+kampanye;
    const totalAdminRasioPct = pctRound(totalAdminKomponen, totalPendapatan);
    const totalAdminRasio=totalPendapatan>0?-(totalAdminKomponen/totalPendapatan)*100:0;
    const amsRasio = amsRasioPct; // alias
    const iklanRasio = totalPendapatan > 0 ? (iklan / totalPendapatan) * 100 : 0; // 2 desimal (23.52%)
    const oprRasio   = totalPendapatan > 0 ? (opr   / totalPendapatan) * 100 : 0; // 2 desimal (24.26%)
    const gpm=totalPendapatan>0?((totalPendapatan-hpp)/totalPendapatan)*100:0;
    const labaRugi=totalPenghasilan-hpp-opr-iklan;
    const npm=totalPendapatan>0?(labaRugi/totalPendapatan)*100:0;

    // ── Target & Minimum ──
    // GPM Minimum = semua beban non-HPP / pendapatan (BEP: laba = 0)
    // GPM min agar BEP: HPP/pend + admin%/pend + iklan%/pend + opr%/pend = 1  →  gpmMin = (admin+iklan+opr)/pend*100 + (hpp/pend)*100... 
    // totalBebanNonHpp untuk GPM Min: admin (tanpa isiSaldo) + iklan + opr
    // IsiSaldo sudah termasuk dalam komponen iklan (real cost), tidak dihitung terpisah
    const totalBebanNonHpp = totalAdminKomponen + iklan + opr;
    const gpmMin = totalPendapatan > 0 ? (totalBebanNonHpp / totalPendapatan) * 100 : 0;
    // ROAS Target BEP: ROAS saat GPM aktual bisa menutup semua beban
    // laba = totalPenghasilan - hpp - opr - iklanTarget = 0  → iklanTarget = totalPenghasilan - hpp - opr
    // ROASmin = totalPendapatan / iklanTarget
    const iklanMaxBep = totalPenghasilan - hpp - opr;
    const roasTargetBep = iklanMaxBep > 0 ? totalPendapatan / iklanMaxBep : 0;
    // ROAS Target NPM 10%: laba = 10% * totalPendapatan  → iklanTarget = totalPenghasilan - hpp - opr - 0.1*totalPendapatan
    const iklanMaxNpm10 = totalPenghasilan - hpp - opr - (totalPendapatan * 0.1);
    const roasTargetNpm10 = iklanMaxNpm10 > 0 ? totalPendapatan / iklanMaxNpm10 : 0;

    // Enable export button jika ada data
    const hasData=totalPendapatan>0||totalPenghasilan>0;
    const btn=document.getElementById('btnExportPDF');
    if(btn){ btn.style.opacity=hasData?'1':'0.4'; btn.style.pointerEvents=hasData?'auto':'none'; }

    // ── Render tabel ──
    // pct: v=nilai persen, cls=class warna, int=true→tampil integer (sesuai spreadsheet)
    const pct=(v,cls,int=false)=>`<span class="rk-pct ${cls}">${v>=0?'+':''}${int?Math.round(v):v.toFixed(2)}%</span>`;
    const rp=(v,cls='')=>`<span class="${cls||''}">${v<0?'('+formatRp(Math.abs(v))+')':formatRp(v)}</span>`;
    const neg=v=>v<0?'rk-val-neg':'rk-val-pos';

    const rows=[
        {type:'normal', label:'TOTAL PENDAPATAN', val:totalPendapatan, pct:null, bold:true},
        {type:'normal', label:'TOTAL PENGHASILAN', val:totalPenghasilan, pct:null, bold:true},
        {type:'normal', label:'HPP', val:-hpp, pct:null, bold:true},
        {type:'normal', label:'OPERASIONAL', val:-opr, pct:oprRasio, pctCls:'rk-pct-yellow', bold:true},
        {type:'normal', label:'IKLAN (real cost)', val:-iklan, pct:iklanRasio, pctCls:'rk-pct-yellow', bold:true},
        {type:'header', label:'RASIO ADMIN DAN LAYANAN'},
        {type:'sub', label:'Biaya Komisi AMS', val:-ams, pct:amsRasioPct, pctCls:'rk-pct-red', pctInt:true},
        {type:'sub', label:'Biaya Administrasi', val:-adminFee, pct:adminFeeRasioPct, pctCls:'rk-pct-red', pctInt:true},
        {type:'sub', label:'Biaya Layanan', val:-layanan, pct:layananRasioPct, pctCls:'rk-pct-red', pctInt:true},
        {type:'sub', label:'Biaya Proses Pesanan', val:-proses, pct:prosesRasioPct, pctCls:'rk-pct-red', pctInt:true},
        {type:'sub', label:'Biaya Kampanye', val:-kampanye, pct:kampanyeRasioPct, pctCls:'rk-pct-red', pctInt:true},
        ...(isiSaldo > 0 ? [{type:'sub', label:'Biaya Isi Saldo Otomatis (dari Penghasilan)', val:-isiSaldo, pct:isiSaldoRasioPct, pctCls:'rk-pct-red', pctInt:true}] : []),
        {type:'header', label:'METRIK KINERJA'},
        {type:'normal', label:'AOV AKTUAL', val:aov, pct:null, bold:true},
        {type:'normal', label:'TOTAL ORDER', val:totalOrder, pct:null, bold:true, isQty:true},
        {type:'normal', label:'ROAS AKTUAL',
            val: null,
            acosText: acosAktual > 0 ? 'ACOS '+acosAktual.toFixed(2)+'%' : '—',
            acosCls: acosAktual > 0 ? (acosAktual<=20?'rk-pct-green':acosAktual<=35?'rk-pct-yellow':'rk-pct-red') : '',
            pctText: roasAktual > 0 ? roasAktual : '—',
            bold: true},
        {type:'header', label:'PROFITABILITAS'},
        {type:'normal', label:'Gros Profit Margin',
            val: totalPendapatan > 0 ? (totalPenghasilan - hpp) : null,
            pctText: totalPendapatan > 0 ? gpm.toFixed(2)+'%' : '—',
            pctCls: totalPendapatan > 0 ? (gpm>=40?'rk-pct-green':gpm>=20?'rk-pct-yellow':'rk-pct-red') : '', bold:true},
        {type:'normal', label:'Net Profit Margin',
            val: totalPendapatan > 0 ? labaRugi : null,
            pctText: totalPendapatan > 0 ? npm.toFixed(2)+'%' : '—',
            pctCls: totalPendapatan > 0 ? (npm>=10?'rk-pct-green':npm>=0?'rk-pct-yellow':'rk-pct-red') : '', bold:true},
        {type:'laba', label:'LABA / RUGI', val:labaRugi},
    ];

    let html='';
    rows.forEach(r=>{
        if(r.type==='header'){
            if(r.label==='RASIO ADMIN DAN LAYANAN'){
                const tAdmV = totalAdminKomponen>0?'('+formatRp(totalAdminKomponen)+')':'—';
                const tAdmP = totalAdminRasioPct!==0?`<span class="rk-admin-badge">⚠ ${totalAdminRasioPct}%</span>`:'';
                html+=`<tr class="rk-row-admin-total">
                    <td>
                        TOTAL BIAYA ADMIN &amp; LAYANAN
                        <span style="font-size:0.68em;font-weight:600;color:#6b7280;letter-spacing:0.3px;margin-left:6px;">(base harga)</span>
                    </td>
                    <td style="text-align:right;font-size:1em;color:#1a1a1a;font-weight:900;letter-spacing:0.3px;">${tAdmV}</td>
                    <td style="text-align:center;padding:8px 10px;">${tAdmP}</td>
                </tr>`;
            } else {
                html+=`<tr class="rk-row-header"><td colspan="3">${r.label}</td></tr>`;
            }
        } else if(r.type==='laba'){
            const isPos=r.val>=0;
            html+=`<tr class="rk-row-laba" style="background:${isPos?'#f0fdf4':'#fff0f0'};border-top:2px solid ${isPos?'#22c55e':'#ef4444'};">
                <td style="font-weight:800;color:${isPos?'#14532d':'#1a1a1a'};">${r.label}</td>
                <td colspan="2" style="text-align:right;font-size:1.1em;font-weight:800;color:${isPos?'#14532d':'#1a1a1a'};">${isPos?formatRp(r.val):'('+formatRp(Math.abs(r.val))+')'}</td>
            </tr>`;
        } else if(r.type==='sub'){
            const vDisplay=r.val!==undefined&&r.val!==null?(r.val<0?'('+formatRp(Math.abs(r.val))+')':formatRp(r.val)):'—';
            const pDisplay=r.pct!==undefined&&r.pct!==null?pct(r.pct,r.pctCls||'',r.pctInt||false):'';
            const subPct = r.pct!==undefined&&r.pct!==null ? `<span class="rk-pct ${r.pctCls||''}" style="font-size:0.82em;font-weight:800;min-width:58px;text-align:center;display:inline-block;padding:3px 8px;">${(r.pct>=0?'+':'')+r.pct.toFixed(2)}%</span>` : '';
            html+=`<tr class="rk-row-sub">
                <td style="color:#666;">↳ ${r.label}</td>
                <td style="text-align:right;font-weight:700;color:${r.val<0?'#991b1b':'#333'};">${vDisplay}</td>
                <td style="text-align:center;padding:6px 10px;">${subPct}</td>
            </tr>`;
        } else {
            let vDisplay='—', pDisplay='';
            if(r.isQty){ vDisplay=`<strong>${r.val.toLocaleString('id-ID')}</strong>`; }
            else if(r.acosText!==undefined){ vDisplay=`<span class="rk-pct ${r.acosCls||''}" style="font-size:0.82em;font-weight:800;padding:3px 9px;">${r.acosText}</span>`; }
            else if(r.val!==null&&r.val!==undefined){ vDisplay=`<strong>${r.val<0?'('+formatRp(Math.abs(r.val))+')':formatRp(r.val)}</strong>`; }
            const valColor = (r.val!==null&&r.val!==undefined&&r.val<0)?'#991b1b':'#333';
            if(r.pctText!==undefined){ pDisplay=`<span class="rk-pct ${r.pctCls||''}" style="font-size:0.9em;font-weight:800;min-width:65px;text-align:center;display:inline-block;padding:4px 10px;">${r.pctText}</span>`; }
            else if(r.pct!==undefined&&r.pct!==null){ pDisplay=pct(r.pct,r.pctCls||'',r.pctInt||false); }
            html+=`<tr class="rk-row-normal">
                <td style="${r.bold?'font-weight:700;':''}color:#222;">${r.label}</td>
                <td style="text-align:right;color:${valColor};">${vDisplay}</td>
                <td style="text-align:center;padding:8px 10px;">${pDisplay}</td>
            </tr>`;
        }
    });

    document.getElementById('rk_tbody').innerHTML=html;

    // Sync OPR dari master data jika belum diisi manual
    if(!document.getElementById('rk_oprTotal').value && localStorage.getItem('oprTotalMonth')){
        document.getElementById('rk_oprTotal').value=localStorage.getItem('oprTotalMonth');
        formatInputRibuan(document.getElementById('rk_oprTotal'));
    }

    // Render grafik order harian
    renderGrafikHarian();
    // ResizeObserver: pastikan canvas selalu mentok ke kanan mengikuti container
    const _chartCanvas = document.getElementById('chartOrderHarian');
    if (_chartCanvas && !_chartCanvas._resizeObserver) {
        const _ro = new ResizeObserver(() => renderGrafikHarian());
        _ro.observe(_chartCanvas.parentElement);
        _chartCanvas._resizeObserver = _ro;
    }

    // Render 4 widget baru scaleup
    renderBepVolume(totalPendapatan, totalPenghasilan, hpp, opr, iklan, totalOrder);
    renderProyeksiLaba(totalPendapatan, totalPenghasilan, hpp, opr, iklan, totalOrder);
    renderTargetOrder(totalPendapatan, totalPenghasilan, hpp, opr, iklan, totalOrder, maxOrderBln);
    // Deteksi beban & pembatalan (auto update setiap kali data berubah)
    renderOngkirPembatalan();
    renderGpmRoasCard();

    // Status Keuangan Badge — gunakan npm & gpm yang sudah dihitung di atas
    updateStatusKeuangan(npm, gpm, totalPendapatan);

    updateKonversi();

    // Sync status ke panel Upload & Data
    syncStatusDetail();
}

// ── RENDER GPM MIN & ROAS TARGET CARD (di bawah Varian Terlaris) ─────────
function renderGpmRoasCard() {
    const el = document.getElementById('gpmRoasContent');
    if (!el) return;

    const get = id => smartParseNumber(document.getElementById(id).value);
    const _manualPendapatan  = get('rk_totalPendapatan');
    const _manualPenghasilan = get('rk_totalPenghasilan');
    const _manualHpp         = get('rk_hppTotal');
    const _manualOpr         = get('rk_oprTotal');

    const pendapatan  = _manualPendapatan  > 0 ? _manualPendapatan  : (rkData.income?.totalPendapatan  || 0);
    const penghasilan = _manualPenghasilan > 0 ? _manualPenghasilan : (rkData.income?.totalPenghasilan || 0);
    const hpp         = _manualHpp         > 0 ? _manualHpp         : ((rkData.order1?.totalHpp || 0) + (rkData.order2?.totalHpp || 0));
    const opr         = _manualOpr         > 0 ? _manualOpr         : (smartParseNumber(localStorage.getItem('oprTotalMonth') || '0'));
    const iklan       = rkData.ads?.totalAds || 0;
    const totalOrder  = rkData.order1?.totalOrder || 0;

    if (pendapatan <= 0) {
        el.innerHTML = '<div style="text-align:center;padding:10px;color:#bbb;font-size:0.78em;">Upload data untuk melihat GPM &amp; ROAS</div>';
        return;
    }

    const totalAdminKomp = rkData.income ? (
        (rkData.income.adminBreakdown?.ams     || 0) +
        (rkData.income.adminBreakdown?.admin   || 0) +
        (rkData.income.adminBreakdown?.layanan || 0) +
        (rkData.income.adminBreakdown?.proses  || 0) +
        (rkData.income.adminBreakdown?.kampanye|| 0) +
        (rkData.income.isiSaldo                || 0)
    ) : 0;

    const gpmAktual = pendapatan > 0 ? ((pendapatan - hpp) / pendapatan) * 100 : 0;
    const gpmMin    = pendapatan > 0 ? ((totalAdminKomp + iklan + opr) / pendapatan) * 100 : 0;
    const gpmOk     = gpmAktual >= gpmMin && gpmMin > 0;

    // ROAS Target kalkulasi
    const sisaBep   = pendapatan - hpp - opr - totalAdminKomp;
    const sisaNpm10 = pendapatan - hpp - opr - totalAdminKomp - (pendapatan * 0.1);
    const roasBep   = sisaBep   > 0 ? pendapatan / sisaBep   : 0;
    const roasNpm10 = sisaNpm10 > 0 ? pendapatan / sisaNpm10 : 0;
    const roasAkt   = iklan > 0 ? pendapatan / iklan : 0;
    const roasBepOk = roasAkt > 0 && roasBep  > 0 && roasAkt >= roasBep;
    const roasN10Ok = roasAkt > 0 && roasNpm10 > 0 && roasAkt >= roasNpm10;

    // ── Simulasi: Jika ROAS dinaikkan +1 / +2 poin, GPM yang dikejar? ──
    // Logika: jika ROAS naik dari aktual → beban iklan/pendapatan turun
    // Beban iklan baru = pendapatan / (roasAkt + delta)
    // GPM min baru = (totalAdminKomp + iklanBaru + opr) / pendapatan * 100
    const simRows = roasAkt > 0 ? [1, 2].map(delta => {
        const roasBaru   = roasAkt + delta;
        const iklanBaru  = pendapatan / roasBaru;
        const gpmMinBaru = ((totalAdminKomp + iklanBaru + opr) / pendapatan) * 100;
        const selisih    = gpmMin - gpmMinBaru;
        const ok         = gpmAktual >= gpmMinBaru;
        return `<div style="background:${ok?'#f0fdf4':'#fffbeb'};border:1px solid ${ok?'#86efac':'#fde68a'};border-radius:7px;padding:6px 9px;display:flex;align-items:center;justify-content:space-between;">
            <div>
                <div style="font-size:0.62em;font-weight:800;color:#555;">Jika ROAS +${delta} → <b style="color:#0d9488;">${roasBaru.toFixed(1)}x</b></div>
                <div style="font-size:0.58em;color:#888;margin-top:1px;">GPM min turun dari ${gpmMin.toFixed(1)}% → <b>${gpmMinBaru.toFixed(1)}%</b> <span style="color:#16a34a;">(hemat ${selisih.toFixed(1)}%)</span></div>
            </div>
            <div style="text-align:right;">
                <div style="font-size:0.75em;font-weight:800;color:${ok?'#16a34a':'#d97706'};">${gpmMinBaru.toFixed(1)}%</div>
                <div style="font-size:0.55em;color:#aaa;">GPM dikejar</div>
            </div>
        </div>`;
    }).join('') : '';

    el.innerHTML = `
        <!-- GPM MIN -->
        <div style="background:#f0f9ff;border:1.5px solid #bae6fd;border-radius:8px;padding:8px 10px;">
            <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:5px;">
                <span style="font-size:0.63em;font-weight:800;color:#0369a1;text-transform:uppercase;">📐 GPM Min (BEP)</span>
                <span style="font-size:0.75em;font-weight:800;padding:2px 7px;border-radius:5px;background:${gpmOk?'#dcfce7':'#fee2e2'};color:${gpmOk?'#16a34a':'#dc2626'};">${gpmMin.toFixed(1)}%</span>
            </div>
            <div style="display:flex;align-items:center;gap:6px;">
                <div style="flex:1;background:#fff;border-radius:5px;padding:4px 6px;text-align:center;border:1px solid #e0f2fe;">
                    <div style="font-size:0.55em;color:#888;font-weight:700;margin-bottom:1px;">AKTUAL</div>
                    <div style="font-size:0.88em;font-weight:800;color:${gpmOk?'#16a34a':'#dc2626'};">${gpmAktual.toFixed(1)}%</div>
                </div>
                <div style="font-size:0.65em;color:#94a3b8;">vs</div>
                <div style="flex:1;background:#fff;border-radius:5px;padding:4px 6px;text-align:center;border:1px solid #e0f2fe;">
                    <div style="font-size:0.55em;color:#888;font-weight:700;margin-bottom:1px;">MIN BEP</div>
                    <div style="font-size:0.88em;font-weight:800;color:#0369a1;">${gpmMin.toFixed(1)}%</div>
                </div>
            </div>
        </div>

        <!-- ROAS TARGET -->
        <div style="background:#f5f3ff;border:1.5px solid #ddd6fe;border-radius:8px;padding:8px 10px;">
            <div style="font-size:0.63em;font-weight:800;color:#7c3aed;text-transform:uppercase;margin-bottom:5px;">🎯 ROAS Target</div>
            <div style="display:flex;gap:4px;">
                <div style="flex:1;background:#fff;border-radius:5px;padding:5px 5px;text-align:center;border:1px solid #ede9fe;">
                    <div style="font-size:0.52em;font-weight:700;color:#7c3aed;margin-bottom:1px;">BEP</div>
                    <div style="font-size:0.85em;font-weight:800;color:${roasBepOk?'#16a34a':'#dc2626'};">${roasBep>0?roasBep.toFixed(1):'N/A'}</div>
                </div>
                <div style="flex:1;background:#fff;border-radius:5px;padding:5px 5px;text-align:center;border:1px solid #ede9fe;">
                    <div style="font-size:0.52em;font-weight:700;color:#7c3aed;margin-bottom:1px;">NPM10%</div>
                    <div style="font-size:0.85em;font-weight:800;color:${roasN10Ok?'#16a34a':'#dc2626'};">${roasNpm10>0?roasNpm10.toFixed(1):'N/A'}</div>
                </div>
                <div style="flex:1;background:#fff;border-radius:5px;padding:5px 5px;text-align:center;border:1px solid #ede9fe;">
                    <div style="font-size:0.52em;font-weight:700;color:#555;margin-bottom:1px;">Aktual</div>
                    <div style="font-size:0.85em;font-weight:800;color:${roasBepOk?'#16a34a':'#dc2626'};">${roasAkt>0?roasAkt.toFixed(1):'—'}</div>
                </div>
            </div>
            ${(roasBep<=0||roasNpm10<=0) ? `<div style="font-size:0.58em;color:#a78bfa;margin-top:5px;line-height:1.4;background:#ede9fe;border-radius:5px;padding:4px 7px;">⚠️ N/A = HPP+OPR+Admin sudah melebihi penghasilan — perlu naikkan harga atau turunkan beban dulu</div>` : ''}
        </div>

        <!-- SIMULASI NAIK ROAS -->
        ${roasAkt > 0 ? `
        <div style="border-top:1px solid #f0f0f0;padding-top:6px;">
            <div style="font-size:0.6em;font-weight:800;color:#555;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:5px;">📈 Simulasi Naik ROAS — GPM yang Dikejar:</div>
            <div style="display:flex;flex-direction:column;gap:4px;">${simRows}</div>
        </div>` : ''}
    `;
}

// ── STATUS KEUANGAN BISNIS ─────────────────────────────────────
function updateStatusKeuangan(npm, gpm, pendapatan) {
    const badge = document.getElementById('statusKeuanganBadge');
    if (!badge) return;
    if (pendapatan <= 0) { badge.style.display = 'none'; return; }

    let label, emoji, bg, color, border;
    if (npm >= 15) {
        label='PROFIT SOLID'; emoji='🟢'; bg='#dcfce7'; color='#15803d'; border='#86efac';
    } else if (npm >= 5) {
        label='SEHAT'; emoji='🟢'; bg='#f0fdf4'; color='#16a34a'; border='#bbf7d0';
    } else if (npm >= 0) {
        label='TIPIS'; emoji='🟡'; bg='#fefce8'; color='#a16207'; border='#fde047';
    } else if (npm >= -10) {
        label='WASPADA'; emoji='🟠'; bg='#fff7ed'; color='#c2410c'; border='#fed7aa';
    } else if (npm >= -25) {
        label='BLEEDING'; emoji='🔴'; bg='#fff0f0'; color='#dc2626'; border='#fca5a5';
    } else {
        label='KRITIS'; emoji='🚨'; bg='#fee2e2'; color='#991b1b'; border='#f87171';
    }

    badge.style.cssText = `display:flex;align-items:center;gap:6px;background:${bg};border:1.5px solid ${border};border-radius:20px;padding:5px 14px;font-size:0.75em;font-weight:800;color:${color};letter-spacing:0.3px;`;
    badge.innerHTML = `${emoji} ${label} <span style="font-weight:500;opacity:0.75;font-size:0.9em;">NPM ${npm.toFixed(1)}%</span>`;
}

// ── RENDER ONGKIR SELLER & PEMBATALAN (Analisis Produk) ───────
function renderOngkirPembatalan() {
    const box = document.getElementById('ongkirBatalBox');
    if (!box) return;
    const d        = rkData.ongkirBatal || {};
    const pendapatan = rkData.income?.totalPendapatan || 0;
    const totalOrder = rkData.order1?.totalOrder || 0;

    // ── DATA ONGKIR SELLER (dari Income) ──
    const ongkir      = d.totalOngkirSeller || 0;
    const pctOngkir   = pendapatan > 0 ? (ongkir / pendapatan * 100) : 0;
    const hasIncome   = !!rkData.income;

    // ── DATA PEMBATALAN (dari parentskudetail — lebih akurat) ──
    const batalPerforma   = d.totalBatalPerforma  !== undefined ? d.totalBatalPerforma  : null;
    const pctBatalPerf    = d.pctBatalPerforma    !== undefined ? d.pctBatalPerforma    : null;
    const revHilang       = d.revenueHilang       || 0;
    const aovPerf         = d.aovPerforma         || 0;
    const orderDibuat     = d.totalOrderDibuat    || 0;
    const orderSiap       = d.totalOrderSiap      || 0;
    // Fallback: dari order sheet jika performa belum upload
    const batalOrder      = d.totalBatal          || 0;
    const batalFinal      = batalPerforma !== null ? batalPerforma : batalOrder;
    const pctBatalFinal   = batalPerforma !== null ? pctBatalPerf  :
                            (totalOrder+batalOrder)>0 ? (batalOrder/(totalOrder+batalOrder)*100) : 0;
    const sumberBatal     = batalPerforma !== null ? 'dari Performa Produk' : 'dari file Pesanan';

    // ── DATA RETUR (dari Income) ──
    const returDana    = d.totalReturDana   || 0;
    const returOngkir  = d.totalReturOngkir || 0;
    const totalRetur   = returDana + returOngkir;
    // Fallback: retur qty dari order sheet
    const returQty     = d.totalRetur       || 0;

    // Tidak ada data sama sekali
    if (!hasIncome && batalFinal === 0 && totalRetur === 0) {
        box.style.display = 'none';
        return;
    }
    box.style.display = 'block';

    // ── SEVERITY ──
    const sevOngkir = pctOngkir > 5 ? 'red' : pctOngkir > 2 ? 'orange' : 'ok';
    const sevBatal  = pctBatalFinal > 10 ? 'red' : pctBatalFinal > 5 ? 'orange' : 'ok';
    const sevRetur  = (returDana > 0 || returQty > 0) ? 'orange' : 'ok';
    const anyRed    = sevOngkir==='red' || sevBatal==='red';
    const anyOrange = !anyRed && (sevOngkir==='orange' || sevBatal==='orange' || sevRetur==='orange');

    const boxBg     = anyRed ? 'rgba(220,38,38,0.05)' : anyOrange ? 'rgba(234,88,12,0.05)' : 'rgba(22,163,74,0.04)';
    const boxBorder = anyRed ? '#fca5a5'  : anyOrange ? '#fdba74'  : '#86efac';
    const titleIcon = anyRed ? '🚨' : anyOrange ? '⚠️' : '✅';
    const titleCol  = anyRed ? '#dc2626' : anyOrange ? '#c2410c' : '#15803d';

    // ── HELPER ──
    const rpFmt = v => v >= 1000000 ? 'Rp '+(v/1000000).toFixed(1)+'jt'
                     : v >= 1000    ? 'Rp '+(v/1000).toFixed(0)+'rb'
                     : 'Rp '+v.toLocaleString('id-ID');
    const sev2col = s => s==='red'?'#dc2626':s==='orange'?'#d97706':'#16a34a';
    const sev2bg  = s => s==='red'?'rgba(220,38,38,0.07)':s==='orange'?'rgba(234,88,12,0.07)':'rgba(22,163,74,0.07)';
    const sev2border = s => s==='red'?'#fca5a5':s==='orange'?'#fdba74':'#86efac';

    // ── CARD 1: BEBAN ONGKIR SELLER ──
    const cardOngkir = `
        <div style="background:${sev2bg(sevOngkir)};border:1.5px solid ${sev2border(sevOngkir)};border-radius:10px;padding:12px 14px;">
            <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:8px;">
                <div style="font-size:0.65em;font-weight:800;color:${sev2col(sevOngkir)};text-transform:uppercase;letter-spacing:0.4px;">📦 Beban Ongkir Seller</div>
                <div style="font-size:1.2em;font-weight:800;color:${sev2col(sevOngkir)};">${pctOngkir>0?pctOngkir.toFixed(2)+'%':'0%'}</div>
            </div>
            <div style="display:grid;grid-template-columns:1fr 1fr;gap:6px;margin-bottom:7px;">
                <div style="background:rgba(255,255,255,0.7);border-radius:7px;padding:6px 8px;text-align:center;">
                    <div style="font-size:0.55em;color:#888;font-weight:700;text-transform:uppercase;margin-bottom:2px;">Total Beban</div>
                    <div style="font-size:0.85em;font-weight:800;color:${sev2col(sevOngkir)};">${ongkir>0?rpFmt(ongkir):'Rp 0'}</div>
                </div>
                <div style="background:rgba(255,255,255,0.7);border-radius:7px;padding:6px 8px;text-align:center;">
                    <div style="font-size:0.55em;color:#888;font-weight:700;text-transform:uppercase;margin-bottom:2px;">Dari Revenue</div>
                    <div style="font-size:0.85em;font-weight:800;color:#555;">${pendapatan>0?rpFmt(pendapatan):'—'}</div>
                </div>
            </div>
            <div style="font-size:0.6em;padding:4px 7px;border-radius:5px;background:rgba(255,255,255,0.6);color:${sev2col(sevOngkir)};font-weight:700;">
                ${ongkir===0?'✅ Tidak ada beban ongkir seller bulan ini':
                  pctOngkir>5?'🚨 Besar — evaluasi program gratis ongkir seller':
                  pctOngkir>2?'⚠️ Perlu dipantau — pertimbangkan batasi promo':
                  '✅ Normal — dalam batas wajar'}
            </div>
            ${!hasIncome?'<div style="font-size:0.58em;color:#aaa;margin-top:4px;font-style:italic;">Upload Income untuk deteksi otomatis</div>':''}
        </div>`;

    // ── CARD 2: PEMBATALAN ──
    const cardBatal = `
        <div style="background:${sev2bg(sevBatal)};border:1.5px solid ${sev2border(sevBatal)};border-radius:10px;padding:12px 14px;">
            <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:8px;">
                <div style="font-size:0.65em;font-weight:800;color:${sev2col(sevBatal)};text-transform:uppercase;letter-spacing:0.4px;">❌ Pembatalan Pesanan</div>
                <div style="font-size:1.2em;font-weight:800;color:${sev2col(sevBatal)};">${pctBatalFinal>0?pctBatalFinal.toFixed(1)+'%':'0%'}</div>
            </div>
            <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:5px;margin-bottom:7px;">
                <div style="background:rgba(255,255,255,0.7);border-radius:7px;padding:5px 6px;text-align:center;">
                    <div style="font-size:0.52em;color:#888;font-weight:700;text-transform:uppercase;margin-bottom:1px;">Dibuat</div>
                    <div style="font-size:0.9em;font-weight:800;color:#333;">${orderDibuat>0?orderDibuat:totalOrder+batalFinal}</div>
                    <div style="font-size:0.5em;color:#aaa;">pesanan</div>
                </div>
                <div style="background:rgba(255,255,255,0.7);border-radius:7px;padding:5px 6px;text-align:center;">
                    <div style="font-size:0.52em;color:#888;font-weight:700;text-transform:uppercase;margin-bottom:1px;">Batal</div>
                    <div style="font-size:0.9em;font-weight:800;color:${sev2col(sevBatal)};">${batalFinal}</div>
                    <div style="font-size:0.5em;color:#aaa;">pesanan</div>
                </div>
                <div style="background:rgba(255,255,255,0.7);border-radius:7px;padding:5px 6px;text-align:center;">
                    <div style="font-size:0.52em;color:#888;font-weight:700;text-transform:uppercase;margin-bottom:1px;">Rev Hilang</div>
                    <div style="font-size:0.75em;font-weight:800;color:${sev2col(sevBatal)};">${revHilang>0?rpFmt(revHilang):'—'}</div>
                </div>
            </div>
            ${revHilang>0?`<div style="font-size:0.58em;color:#888;margin-bottom:5px;">Est. hilang = ${batalFinal} × AOV ${rpFmt(aovPerf)} — <i>${sumberBatal}</i></div>`:''}
            <div style="font-size:0.6em;padding:4px 7px;border-radius:5px;background:rgba(255,255,255,0.6);color:${sev2col(sevBatal)};font-weight:700;">
                ${batalFinal===0?'✅ Tidak ada pembatalan terdeteksi':
                  pctBatalFinal>10?'🚨 Tinggi — cek stok, penipuan, & respons toko':
                  pctBatalFinal>5?'⚠️ Perlu perhatian — monitor penyebab pembatalan':
                  '⚠️ Dalam batas — tetap monitor'}
            </div>
        </div>`;

    // ── CARD 3: RETUR / SIAP KIRIM BALIK ──
    const hasRetur = totalRetur > 0 || returQty > 0;
    const cardRetur = `
        <div style="background:${hasRetur?'rgba(124,58,237,0.06)':'rgba(22,163,74,0.04)'};border:1.5px solid ${hasRetur?'#ddd6fe':'#86efac'};border-radius:10px;padding:12px 14px;">
            <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:8px;">
                <div style="font-size:0.65em;font-weight:800;color:${hasRetur?'#7c3aed':'#15803d'};text-transform:uppercase;letter-spacing:0.4px;">🔄 Retur / Siap Kirim Balik</div>
                <div style="font-size:1.2em;font-weight:800;color:${hasRetur?'#7c3aed':'#16a34a'};">${returQty>0?returQty+' pcs':hasRetur?'Ada':'—'}</div>
            </div>
            <div style="display:grid;grid-template-columns:1fr 1fr;gap:6px;margin-bottom:7px;">
                <div style="background:rgba(255,255,255,0.7);border-radius:7px;padding:6px 8px;text-align:center;">
                    <div style="font-size:0.52em;color:#888;font-weight:700;text-transform:uppercase;margin-bottom:1px;">Dana Dikembalikan</div>
                    <div style="font-size:0.82em;font-weight:800;color:${returDana>0?'#7c3aed':'#aaa'};">${returDana>0?rpFmt(returDana):'Rp 0'}</div>
                </div>
                <div style="background:rgba(255,255,255,0.7);border-radius:7px;padding:6px 8px;text-align:center;">
                    <div style="font-size:0.52em;color:#888;font-weight:700;text-transform:uppercase;margin-bottom:1px;">Ongkir Retur</div>
                    <div style="font-size:0.82em;font-weight:800;color:${returOngkir>0?'#dc2626':'#aaa'};">${returOngkir>0?rpFmt(returOngkir):'Rp 0'}</div>
                </div>
            </div>
            <div style="font-size:0.6em;padding:4px 7px;border-radius:5px;background:rgba(255,255,255,0.6);color:${hasRetur?'#7c3aed':'#15803d'};font-weight:700;">
                ${!hasRetur?'✅ Tidak ada retur bulan ini':
                  returDana>0?'⚠️ Ada pengembalian dana — cek kualitas & deskripsi produk':
                  '⚠️ Ada item diretur — cek packaging & kualitas produk'}
            </div>
            ${!hasIncome?'<div style="font-size:0.58em;color:#aaa;margin-top:4px;font-style:italic;">Upload Income untuk nilai retur</div>':''}
        </div>`;

    box.innerHTML = `
    <div style="background:${boxBg};border:1.5px solid ${boxBorder};border-radius:14px;padding:14px 16px;">
        <div style="display:flex;align-items:center;gap:8px;margin-bottom:12px;">
            <span style="font-size:1em;">${titleIcon}</span>
            <div style="font-size:0.72em;font-weight:800;color:${titleCol};text-transform:uppercase;letter-spacing:0.5px;">Deteksi Beban & Pembatalan</div>
            <div style="flex:1;"></div>
            <div style="font-size:0.6em;color:#aaa;font-style:italic;">Auto-detect dari Income & Performa Produk</div>
        </div>
        <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px;">
            ${cardOngkir}${cardBatal}${cardRetur}
        </div>
    </div>`;
}


function renderBepVolume(pendapatan, penghasilan, hpp, opr, iklan, totalOrder) {
    const el = document.getElementById('bepVolumeContent');
    if (!el) return;
    if (pendapatan <= 0 || totalOrder <= 0) {
        el.innerHTML = '<div style="text-align:center;padding:10px;color:#bbb;font-size:0.78em;">Upload data untuk melihat BEP</div>';
        return;
    }

    const aov = Math.round(pendapatan / totalOrder);
    // Laba per order sekarang
    const labaPerOrder = (penghasilan - hpp - iklan) / totalOrder;
    // BEP = OPR dibagi laba-per-order-sebelum-OPR
    const labaSebelumOpr = labaPerOrder + (opr / totalOrder);
    const bepOrder = labaSebelumOpr > 0 ? Math.ceil(opr / labaSebelumOpr) : 0;
    const isAboveBep = totalOrder >= bepOrder;
    const selisih = totalOrder - bepOrder;

    const bepColor = isAboveBep ? '#16a34a' : '#dc2626';
    const bepBg = isAboveBep ? '#f0fdf4' : '#fff0f0';
    const bepBorder = isAboveBep ? '#86efac' : '#fca5a5';

    el.innerHTML = `
        <div style="background:${bepBg};border:1px solid ${bepBorder};border-radius:8px;padding:7px 10px;margin-bottom:5px;text-align:center;">
            <div style="font-size:0.58em;font-weight:700;color:#666;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:2px;">BEP Minimum</div>
            <div style="font-size:1.4em;font-weight:800;color:${bepColor};line-height:1;">${bepOrder > 0 ? bepOrder : '—'}</div>
            <div style="font-size:0.58em;color:#666;margin-top:1px;">order / bulan</div>
        </div>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:5px;">
            <div style="background:#f8f9fa;border-radius:7px;padding:5px 8px;text-align:center;">
                <div style="font-size:0.58em;color:#888;font-weight:700;text-transform:uppercase;margin-bottom:1px;">Aktual</div>
                <div style="font-size:1.1em;font-weight:800;color:#1a1a2e;line-height:1.2;">${totalOrder}</div>
                <div style="font-size:0.55em;color:#aaa;">order</div>
            </div>
            <div style="background:${isAboveBep?'#f0fdf4':'#fff0f0'};border-radius:7px;padding:5px 8px;text-align:center;">
                <div style="font-size:0.58em;color:#888;font-weight:700;text-transform:uppercase;margin-bottom:1px;">Selisih</div>
                <div style="font-size:1.1em;font-weight:800;color:${bepColor};line-height:1.2;">${selisih >= 0 ? '+' : ''}${selisih}</div>
                <div style="font-size:0.55em;color:#aaa;">order</div>
            </div>
        </div>
        <div style="font-size:0.62em;color:${bepColor};font-weight:700;text-align:center;margin-top:4px;padding:3px 6px;background:${bepBg};border-radius:6px;line-height:1.4;">
            ${isAboveBep ? '✅ Sudah melewati BEP' : `⚠️ Butuh +${Math.abs(selisih)} order lagi`}
        </div>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:5px;margin-top:5px;">
            <div style="background:#f0f9ff;border:1px solid #bae6fd;border-radius:7px;padding:5px 8px;text-align:center;">
                <div style="font-size:0.55em;color:#0369a1;font-weight:700;text-transform:uppercase;margin-bottom:1px;">BEP / Hari</div>
                <div style="font-size:1em;font-weight:800;color:#0369a1;line-height:1.2;">${bepOrder > 0 ? Math.ceil(bepOrder/30) : '—'}</div>
                <div style="font-size:0.52em;color:#aaa;">order / hari</div>
            </div>
            <div style="background:#f0f9ff;border:1px solid #bae6fd;border-radius:7px;padding:5px 8px;text-align:center;">
                <div style="font-size:0.55em;color:#0369a1;font-weight:700;text-transform:uppercase;margin-bottom:1px;">Aktual / Hari</div>
                <div style="font-size:1em;font-weight:800;color:#0369a1;line-height:1.2;">${(totalOrder/30).toFixed(1)}</div>
                <div style="font-size:0.52em;color:#aaa;">order / hari</div>
            </div>
        </div>`;
}

// ── WIDGET 2: PROYEKSI LABA DI X ORDER ────────────────────────
function renderProyeksiLaba(pendapatan, penghasilan, hpp, opr, iklan, totalOrder) {
    const el = document.getElementById('proyeksiLabaContent');
    if (!el) return;
    if (pendapatan <= 0 || totalOrder <= 0) {
        el.innerHTML = '<div style="text-align:center;padding:10px;color:#bbb;font-size:0.78em;">Upload data untuk melihat proyeksi</div>';
        return;
    }

    const aov = Math.round(pendapatan / totalOrder);
    // Laba per order (tanpa OPR karena OPR fixed)
    const labaPerOrder = (penghasilan - hpp - iklan) / totalOrder;
    
    // Skenario: 1.5x, 2x, 3x dari order sekarang
    const skenario = [
        { label: '1.5×', order: Math.round(totalOrder * 1.5) },
        { label: '2×',   order: Math.round(totalOrder * 2) },
        { label: '3×',   order: Math.round(totalOrder * 3) },
    ];

    const rpShort = v => {
        const abs = Math.abs(v);
        const sign = v < 0 ? '-' : '+';
        if (abs >= 1000000) return sign + 'Rp ' + (abs/1000000).toFixed(1) + 'jt';
        if (abs >= 1000)    return sign + 'Rp ' + (abs/1000).toFixed(0) + 'rb';
        return sign + 'Rp ' + abs.toLocaleString('id-ID');
    };

    const rows = skenario.map(s => {
        const labaProyeksi = (labaPerOrder * s.order) - opr;
        const isProfit = labaProyeksi >= 0;
        const color = isProfit ? '#16a34a' : '#dc2626';
        const bg = isProfit ? '#f0fdf4' : '#fff0f0';
        return `<div style="display:flex;align-items:center;justify-content:space-between;background:${bg};border-radius:7px;padding:6px 9px;">
            <div>
                <span style="font-size:0.68em;font-weight:800;color:#555;text-transform:uppercase;">Jika ${s.label} order</span>
                <div style="font-size:0.62em;color:#aaa;">${s.order} order/bln</div>
            </div>
            <div style="text-align:right;">
                <div style="font-size:0.85em;font-weight:800;color:${color};">${rpShort(labaProyeksi)}</div>
                <div style="font-size:0.6em;color:#aaa;">laba/bln</div>
            </div>
        </div>`;
    }).join('');

    const labaSekarang = (labaPerOrder * totalOrder) - opr;
    const labaColor = labaSekarang >= 0 ? '#16a34a' : '#dc2626';

    el.innerHTML = `
        <div style="font-size:0.65em;font-weight:700;color:#777;margin-bottom:4px;">Sekarang: <span style="color:${labaColor}">${rpShort(labaSekarang)}/bln</span></div>
        <div style="display:flex;flex-direction:column;gap:4px;">${rows}</div>`;
}

// ── WIDGET 3: TARGET ORDER BULAN DEPAN ────────────────────────
function renderTargetOrder(pendapatan, penghasilan, hpp, opr, iklan, totalOrder, maxOrderBln = 0) {
    const el = document.getElementById('targetOrderContent');
    if (!el) return;
    if (pendapatan <= 0 || totalOrder <= 0) {
        el.innerHTML = '<div style="text-align:center;padding:10px;color:#bbb;font-size:0.78em;">Upload data untuk melihat target</div>';
        return;
    }

    const aov = Math.round(pendapatan / totalOrder);
    const labaPerOrder = (penghasilan - hpp - iklan) / totalOrder;
    const orderUntukBEP = labaPerOrder > 0 ? Math.ceil(opr / labaPerOrder) : 0;

    // Target ideal: BEP order + 20% buffer, minimal +10% dari aktual
    const targetIdeal = Math.max(
        Math.round(totalOrder * 1.10),
        orderUntukBEP > 0 ? Math.round(orderUntukBEP * 1.2) : Math.round(totalOrder * 1.1)
    );

    // Dibatasi oleh kapasitas max order jika diisi
    const kapasitasDiisi = maxOrderBln > 0;
    const targetBulanDepan = kapasitasDiisi ? Math.min(targetIdeal, maxOrderBln) : targetIdeal;
    const targetKapasitasTerbatas = kapasitasDiisi && targetIdeal > maxOrderBln;

    const targetOmset = targetBulanDepan * aov;
    const targetLabaProyeksi = (labaPerOrder * targetBulanDepan) - opr;
    const isTargetProfit = targetLabaProyeksi >= 0;

    const rpFmt = v => {
        const abs = Math.abs(v);
        if (abs >= 1000000) return (v<0?'(':'') + 'Rp ' + (abs/1000000).toFixed(1) + 'jt' + (v<0?')':'');
        return (v<0?'(':'') + 'Rp ' + abs.toLocaleString('id-ID') + (v<0?')':'');
    };

    const growthPct = Math.round(((targetBulanDepan - totalOrder) / totalOrder) * 100);

    // Peringatan jika kapasitas jadi pembatas
    const warningHtml = targetKapasitasTerbatas ? `
        <div style="background:#fffbeb;border:1px solid #fde68a;border-radius:7px;padding:6px 9px;margin-top:5px;font-size:0.65em;color:#92400e;line-height:1.5;">
            ⚠️ <b>Kapasitas max ${maxOrderBln.toLocaleString('id-ID')} order/bln tercapai.</b><br>
            BEP butuh ${orderUntukBEP} order — perlu naikkan harga atau turunkan beban dulu.
        </div>` : '';

    el.innerHTML = `
        <div style="background:#f5f3ff;border:1px solid #ddd6fe;border-radius:8px;padding:9px 11px;text-align:center;margin-bottom:6px;">
            <div style="font-size:0.62em;font-weight:700;color:#7c3aed;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:2px;">Target Bulan Depan</div>
            <div style="font-size:1.5em;font-weight:800;color:#5b21b6;line-height:1;">${targetBulanDepan.toLocaleString('id-ID')}</div>
            <div style="font-size:0.62em;color:#7c3aed;">order &nbsp;|&nbsp; tumbuh <b>+${growthPct}%</b></div>
        </div>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:5px;">
            <div style="background:#f8f9fa;border-radius:7px;padding:6px 8px;text-align:center;">
                <div style="font-size:0.58em;color:#888;font-weight:700;margin-bottom:2px;">TARGET OMSET</div>
                <div style="font-size:0.8em;font-weight:800;color:#1a1a2e;">${rpFmt(targetOmset)}</div>
            </div>
            <div style="background:${isTargetProfit?'#f0fdf4':'#fff0f0'};border-radius:7px;padding:6px 8px;text-align:center;">
                <div style="font-size:0.58em;color:#888;font-weight:700;margin-bottom:2px;">PROYEKSI LABA</div>
                <div style="font-size:0.8em;font-weight:800;color:${isTargetProfit?'#16a34a':'#dc2626'};">${rpFmt(targetLabaProyeksi)}</div>
            </div>
        </div>
        <div style="font-size:0.65em;color:#555;margin-top:5px;padding:4px 8px;background:#fafafa;border-radius:6px;text-align:center;">
            Dari <b>${totalOrder}</b> → <b>${targetBulanDepan.toLocaleString('id-ID')}</b> order &nbsp;(+${(targetBulanDepan - totalOrder).toLocaleString('id-ID')} order)
        </div>
        ${warningHtml}`;
}

// ── WIDGET 4: KONVERSI KUNJUNGAN → ORDER (Opsi B — auto dari performa) ─
function updateKonversi() {
    const resultEl = document.getElementById('konversiResult');
    const inputEl  = document.getElementById('rk_kunjungan');
    if (!resultEl) return;

    const totalOrder = rkData.order1?.totalOrder || 0;

    // ── MODE A: Data dari file Performa Produk (auto) ──────────
    if (rkData.performa) {
        const { produkArr, totalKunjungan, totalOrderDibuat, cvrStore, produkAktif } = rkData.performa;

        // Hide input manual — tidak perlu
        const inputWrap = inputEl?.closest('div');
        if (inputEl) inputEl.closest('div').style.display = 'none';

        const cvrColor  = cvrStore >= 3 ? '#16a34a' : cvrStore >= 1.5 ? '#d97706' : '#dc2626';
        const cvrBg     = cvrStore >= 3 ? '#f0fdf4'  : cvrStore >= 1.5 ? '#fffbeb'  : '#fff0f0';
        const cvrBorder = cvrStore >= 3 ? '#86efac'  : cvrStore >= 1.5 ? '#fde68a'  : '#fca5a5';
        const cvrLabel  = cvrStore >= 3 ? '✅ Di atas target' : cvrStore >= 1.5 ? '⚠️ Perlu ditingkatkan' : '🔴 Di bawah benchmark';

        // Top 3 produk dengan kunjungan terbanyak & CVR-nya
        const topProduk = [...produkArr]
            .filter(p => p.kunjungan > 0)
            .sort((a, b) => b.kunjungan - a.kunjungan)
            .slice(0, 3);

        // Identifikasi masalah: kunjungan tinggi tapi CVR rendah = listing/foto problem
        const produkBermasalah = produkArr
            .filter(p => p.kunjungan > 500 && p.cvrDibuat < 1.0)
            .sort((a, b) => b.kunjungan - a.kunjungan)
            .slice(0, 2);

        const rpShort = v => v >= 1000000 ? 'Rp ' + (v/1000000).toFixed(1) + 'jt'
                           : v >= 1000    ? 'Rp ' + (v/1000).toFixed(0) + 'rb'
                           : 'Rp ' + v.toLocaleString('id-ID');

        const skuRows = topProduk.map(p => {
            const c = p.cvrDibuat;
            const cc = c >= 2 ? '#16a34a' : c >= 1 ? '#d97706' : '#dc2626';
            const shortName = p.namaProduk.length > 20
                ? p.namaProduk.substring(0, 20) + '…' : p.namaProduk;
            return `<div style="display:flex;align-items:center;justify-content:space-between;padding:4px 0;border-bottom:1px solid #f5f5f5;">
                <div>
                    <div style="font-size:0.65em;font-weight:700;color:#333;">${p.skuInduk}</div>
                    <div style="font-size:0.58em;color:#999;">${p.kunjungan.toLocaleString('id-ID')} kunjungan</div>
                </div>
                <div style="text-align:right;">
                    <div style="font-size:0.72em;font-weight:800;color:${cc};">CVR ${c.toFixed(2)}%</div>
                    <div style="font-size:0.58em;color:#aaa;">${p.orderDibuat} order</div>
                </div>
            </div>`;
        }).join('');

        const warningRows = produkBermasalah.length > 0
            ? `<div style="margin-top:6px;background:#fff7ed;border:1px solid #fed7aa;border-radius:7px;padding:7px 9px;">
                <div style="font-size:0.62em;font-weight:800;color:#d97706;margin-bottom:4px;">⚠️ BANYAK DILIHAT, SEDIKIT BELI:</div>
                ${produkBermasalah.map(p =>
                    `<div style="font-size:0.62em;color:#92400e;padding:2px 0;">${p.skuInduk}: ${p.kunjungan.toLocaleString('id-ID')} kunjungan → CVR <b>${p.cvrDibuat.toFixed(2)}%</b> — cek foto/deskripsi</div>`
                ).join('')}
              </div>` : '';

        const potensiOrder3pct = Math.round(totalKunjungan * 0.03);
        const gapOrder = potensiOrder3pct - totalOrderDibuat;

        resultEl.innerHTML = `
            <div style="background:${cvrBg};border:1px solid ${cvrBorder};border-radius:8px;padding:8px 10px;text-align:center;margin-bottom:7px;">
                <div style="font-size:0.58em;font-weight:700;color:#555;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:2px;">CVR Toko — Auto dari Performa</div>
                <div style="font-size:1.6em;font-weight:800;color:${cvrColor};line-height:1;">${cvrStore.toFixed(2)}%</div>
                <div style="font-size:0.6em;color:#888;margin-top:2px;">${totalOrderDibuat} order dari ${totalKunjungan.toLocaleString('id-ID')} kunjungan</div>
                <div style="font-size:0.6em;font-weight:700;color:${cvrColor};margin-top:2px;">${cvrLabel}</div>
            </div>

            <div style="display:grid;grid-template-columns:1fr 1fr;gap:4px;margin-bottom:6px;">
                <div style="background:#f8f9fa;border-radius:6px;padding:5px 7px;text-align:center;">
                    <div style="font-size:0.57em;color:#888;font-weight:700;margin-bottom:1px;">Benchmark</div>
                    <div style="font-size:0.85em;font-weight:800;color:#555;">2.0%</div>
                </div>
                <div style="background:#f0f9ff;border-radius:6px;padding:5px 7px;text-align:center;">
                    <div style="font-size:0.57em;color:#888;font-weight:700;margin-bottom:1px;">Target Ideal</div>
                    <div style="font-size:0.85em;font-weight:800;color:#d97706;">3.0%</div>
                </div>
            </div>

            ${gapOrder > 0 ? `<div style="font-size:0.63em;color:#d97706;font-weight:700;text-align:center;background:#fffbeb;border-radius:6px;padding:4px 7px;margin-bottom:6px;">
                💡 Jika CVR 3% → potensi +${gapOrder} order/bln lagi
            </div>` : `<div style="font-size:0.63em;color:#16a34a;font-weight:700;text-align:center;background:#f0fdf4;border-radius:6px;padding:4px 7px;margin-bottom:6px;">
                ✅ CVR di atas 3% — excellent!
            </div>`}

            <div style="font-size:0.62em;font-weight:800;color:#555;text-transform:uppercase;margin-bottom:4px;">Top Produk by Kunjungan:</div>
            ${skuRows}
            ${warningRows}`;

        return;
    }

    // ── MODE B: Fallback input manual (jika performa belum diupload) ─
    if (inputEl) inputEl.closest('div').style.display = 'block';
    const kunjungan = inputEl ? smartParseNumber(inputEl.value) : 0;

    if (totalOrder <= 0) {
        resultEl.innerHTML = '<span style="color:#bbb;font-size:0.75em;">Upload data pesanan dulu</span>';
        return;
    }
    if (kunjungan <= 0) {
        resultEl.innerHTML = `<div style="color:#bbb;font-size:0.75em;text-align:center;line-height:1.6;">
            Isi kunjungan manual di atas,<br>atau upload file<br><b style="color:#d97706;">Performa Produk</b><br>untuk auto-hitung CVR 📊</div>`;
        return;
    }

    const cvr = (totalOrder / kunjungan) * 100;
    const cvrColor  = cvr >= 3 ? '#16a34a' : cvr >= 1.5 ? '#d97706' : '#dc2626';
    const cvrBg     = cvr >= 3 ? '#f0fdf4'  : cvr >= 1.5 ? '#fffbeb'  : '#fff0f0';
    const cvrBorder = cvr >= 3 ? '#86efac'  : cvr >= 1.5 ? '#fde68a'  : '#fca5a5';
    const tambahOrder = Math.round(kunjungan * 0.03) - totalOrder;

    resultEl.innerHTML = `
        <div style="background:${cvrBg};border:1px solid ${cvrBorder};border-radius:8px;padding:8px 10px;text-align:center;margin-bottom:6px;">
            <div style="font-size:0.62em;font-weight:700;color:#555;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:2px;">Conversion Rate</div>
            <div style="font-size:1.6em;font-weight:800;color:${cvrColor};line-height:1;">${cvr.toFixed(2)}%</div>
            <div style="font-size:0.6em;color:#888;margin-top:2px;">${totalOrder} order dari ${kunjungan.toLocaleString('id-ID')} kunjungan</div>
        </div>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:4px;margin-bottom:5px;">
            <div style="background:#f8f9fa;border-radius:6px;padding:5px 7px;text-align:center;">
                <div style="font-size:0.58em;color:#888;font-weight:700;margin-bottom:1px;">Benchmark</div>
                <div style="font-size:0.85em;font-weight:800;color:#555;">2.0%</div>
            </div>
            <div style="background:#fffbeb;border-radius:6px;padding:5px 7px;text-align:center;">
                <div style="font-size:0.58em;color:#888;font-weight:700;margin-bottom:1px;">Target Ideal</div>
                <div style="font-size:0.85em;font-weight:800;color:#d97706;">3.0%</div>
            </div>
        </div>
        ${tambahOrder > 0
            ? `<div style="font-size:0.65em;color:#d97706;font-weight:700;text-align:center;background:#fffbeb;border-radius:6px;padding:4px 7px;">
                💡 Jika CVR 3% → potensi +${tambahOrder} order/bln lagi</div>`
            : `<div style="font-size:0.65em;color:#16a34a;font-weight:700;text-align:center;background:#f0fdf4;border-radius:6px;padding:4px 7px;">
                ✅ CVR di atas target — pertahankan!</div>`
        }`;
}

// ── EXPORT PDF ────────────────────────────────────────────────
function exportRasioPDF() {
    const periode=document.getElementById('rk_periode').value||'Periode';
    const table=document.getElementById('rk_table');
    if(!table) return;

    const win=window.open('','_blank');
    win.document.write(`<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Rasio Keuangan ${periode}</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;600;700;800&display=swap');
  *{box-sizing:border-box;margin:0;padding:0;}
  body{font-family:'Plus+Jakarta+Sans',Arial,sans-serif;background:#fff;padding:24px;color:#1a1a2e;font-size:12px;}
  .header{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:20px;padding-bottom:14px;border-bottom:3px solid #ee4d2d;}
  .header-left h1{font-size:16px;font-weight:800;color:#ee4d2d;margin-bottom:2px;}
  .header-left p{font-size:10px;color:#888;font-weight:500;}
  .header-right{text-align:right;}
  .header-right .bln{font-size:20px;font-weight:800;color:#1a1a2e;}
  .header-right .sub{font-size:9px;color:#aaa;text-transform:uppercase;letter-spacing:1px;}
  table{width:100%;border-collapse:collapse;font-size:11.5px;}
  th{background:#ee4d2d;color:#fff;padding:8px 12px;text-align:left;font-weight:800;font-size:11px;letter-spacing:0.5px;}
  th:last-child,th:nth-child(2){text-align:right;}
  .rk-row-header td{background:#fef9c3;color:#92400e;font-weight:800;font-size:10px;text-transform:uppercase;padding:6px 12px;letter-spacing:0.5px;}
  .rk-row-normal td{padding:7px 12px;border-bottom:1px solid #f0f0f0;}
  .rk-row-sub td{padding:5px 12px 5px 22px;border-bottom:1px solid #f9f9f9;color:#666;font-size:11px;}
  .rk-row-laba td{padding:9px 12px;font-weight:800;font-size:13px;}
  .val-right{text-align:right;}
  .badge{display:inline-block;font-size:10px;font-weight:700;padding:1px 6px;border-radius:4px;}
  .badge-red{background:#fee2e2;color:#991b1b;}
  .badge-yellow{background:#fef9c3;color:#854d0e;}
  .badge-green{background:#dcfce7;color:#166534;}
  .footer{margin-top:20px;padding-top:10px;border-top:1px solid #eee;font-size:9px;color:#bbb;text-align:center;}
  @media print{body{padding:0;} .no-print{display:none;}}
</style>
</head>
<body>
<div class="header">
  <div class="header-left">
    <h1>📊 LAPORAN RASIO KEUANGAN</h1>
    <p>Shopee Price Projection — Burhanmology</p>
  </div>
  <div class="header-right">
    <div class="bln">${periode}</div>
    <div class="sub">Laporan Bulanan</div>
  </div>
</div>
${table.outerHTML}
<div class="footer">
  Dicetak pada ${new Date().toLocaleDateString('id-ID',{weekday:'long',year:'numeric',month:'long',day:'numeric'})} &nbsp;|&nbsp; Shopee Price Projection by Burhanmology
</div>
<div class="no-print" style="margin-top:20px;text-align:center;">
  <button onclick="window.print()" style="background:#ee4d2d;color:#fff;border:none;padding:10px 24px;border-radius:8px;font-size:13px;font-weight:700;cursor:pointer;">🖨️ Print / Save as PDF</button>
</div>
</body>
</html>`);
    win.document.close();
    setTimeout(()=>win.print(), 800);
}

// ── RESET ─────────────────────────────────────────────────────
function resetRasio() {
    rkData={income:null,order1:null,order2:null,ads:null,performa:null,periode:'',_incomeNoPesanan:new Set()};
    ['rk_totalPendapatan','rk_totalPenghasilan','rk_hppTotal','rk_oprTotal','rk_adminTotal','rk_amsTotal','rk_periode','adsManualVal'].forEach(id=>{
        const el=document.getElementById(id); if(el) el.value='';
    });
    ['boxIncome','boxOrder1','boxOrder2'].forEach(id=>{
        const b=document.getElementById(id); if(b) b.classList.remove('uploaded');
    });
    ['statusIncome','statusOrder1','statusOrder2','statusAds'].forEach(id=>{
        const el=document.getElementById(id); if(el) el.innerText='Belum upload';
    });
    ['rk_st_income','rk_st_order1','rk_st_order2','rk_st_ads','rk_st_hpp'].forEach(id=>{
        const el=document.getElementById(id); if(el){el.innerText='—';el.style.color='#bbb';}
    });
    ['fileIncome','fileOrder1','fileOrder2','fileAds'].forEach(id=>{
        const el=document.getElementById(id); if(el) el.value='';
    });
    document.getElementById('adsManualWrap').style.display='none';
    document.getElementById('rk_thPeriode').innerText='—';
    document.getElementById('rk_tbody').innerHTML='<tr><td colspan="3" style="text-align:center;padding:30px;color:#bbb;font-size:0.88em;">Upload data untuk melihat dashboard rasio keuangan</td></tr>';
    const btn=document.getElementById('btnExportPDF');
    if(btn){btn.style.opacity='0.4';btn.style.pointerEvents='none';}
    // Reset widget scaleup baru
    const resetMsg = (id, msg) => { const el=document.getElementById(id); if(el) el.innerHTML=`<div style="text-align:center;padding:10px;color:#bbb;font-size:0.78em;">${msg}</div>`; };
    resetMsg('bepVolumeContent','Upload data untuk melihat BEP');
    resetMsg('proyeksiLabaContent','Upload data untuk melihat proyeksi');
    resetMsg('targetOrderContent','Upload data untuk melihat target');
    const kEl=document.getElementById('rk_kunjungan'); if(kEl) kEl.value='';
    const rEl=document.getElementById('konversiResult'); if(rEl) rEl.innerHTML='<span style="color:#bbb;font-size:0.75em;">Isi jumlah kunjungan di atas</span>';
    // Reset performa
    // Hapus nama file yang tampil
    document.querySelectorAll('.rk-upload-filename').forEach(el => el.textContent = '');
    const bPerforma=document.getElementById('boxPerforma'); if(bPerforma) bPerforma.classList.remove('uploaded');
    document.getElementById('statusPerforma').innerText='Belum upload';
    const stPEl=document.getElementById('rk_st_performa_detail'); if(stPEl){stPEl.innerText='—';stPEl.style.color='#bbb';}
    const fPEl=document.getElementById('filePerforma'); if(fPEl) fPEl.value='';
}
// ── Sinkron OPR dari master data saat tab rasio dibuka ───────
window.addEventListener('resize', () => {
    if (typeof renderGrafikHarian === 'function') renderGrafikHarian();
});

document.addEventListener('DOMContentLoaded', function() {

    // ── INIT MULTI-TOKO SESSION ───────────────────────────────
    if (typeof initTokoSession === 'function') {
        initTokoSession().then(() => {
            if (typeof checkSupabaseConnection === 'function') checkSupabaseConnection();
            // Update no-toko state pada switcher
            const sw = document.querySelector('.toko-switcher-inner');
            if (sw) {
                const hasTokoNama = typeof getAktifTokoNama === 'function' && getAktifTokoNama();
                sw.classList.toggle('no-toko', !hasTokoNama);
            }
        });
    }

    // ── TAB RASIO: restore OPR dari localStorage ──────────────
    const origBtns = document.querySelectorAll('.tab-btn');
    origBtns.forEach(btn => {
        const oc = btn.getAttribute('onclick');
        if (oc && oc.includes("'rasio'")) {
            btn.addEventListener('click', function() {
                setTimeout(function(){
                    const oprSaved = localStorage.getItem('oprTotalMonth');
                    const rkOpr = document.getElementById('rk_oprTotal');
                    if (rkOpr && !rkOpr.value && oprSaved) {
                        rkOpr.value = oprSaved;
                        formatInputRibuan(rkOpr);
                        updateRasioDashboard();
                    }
                }, 50);
            });
        }
    });
});

function lookupHppByRef(val) {
    // Dipanggil saat user ketik di field hppRefSku
    const info = document.getElementById('hppRefInfo');
    const hppInput = document.getElementById('hppSatuan');
    const skuActive = document.getElementById('hppRefSkuActive');
    if (!val || val.trim() === '') {
        if (info) info.style.display = 'none';
        if (skuActive) skuActive.style.display = 'none';
        return;
    }
    const found = hppMaster.find(h => h.refSku.toLowerCase().includes(val.toLowerCase().trim()));
    if (found) {
        hppInput.value = found.hpp.toLocaleString('id-ID');
        hppInput.style.borderColor = '#16a34a';
        hppInput.style.background = '#f0fdf4';
        if (info) {
            info.style.display = 'block';
            info.style.color = '#16a34a';
            info.style.fontSize = '0.75em';
            info.style.fontWeight = '600';
            info.innerHTML = '✓ Approved';
        }
        if (skuActive) {
            skuActive.style.display = 'none';
        }
        setTimeout(() => { hppInput.style.borderColor=''; hppInput.style.background=''; }, 2500);
        cleanAndHitung();
    } else {
        hppInput.style.borderColor = '';
        hppInput.style.background = '';
        if (info) {
            if (val && val.trim() !== '') {
                info.style.display = 'block';
                info.style.color = '#dc2626';
                info.style.fontSize = '0.7em';
                info.style.fontStyle = 'italic';
                info.style.fontWeight = '500';
                info.innerHTML = '✗ Data tidak ditemukan, coba update HPP di Master Data';
            } else {
                info.style.display = 'none';
            }
        }
        if (skuActive) skuActive.style.display = 'none';
    }
}

// ══════════════════════════════════════════════
// REKAP TAHUNAN
// ══════════════════════════════════════════════
const REKAP_METRIK = [
    { key:'pendapatan',   label:'TOTAL PENDAPATAN',      type:'rp',      rowClass:'normal bold' },
    { key:'penghasilan',  label:'TOTAL PENGHASILAN',     type:'rp',      rowClass:'normal bold' },
    { key:'hpp',          label:'HPP',                   type:'rp-neg',  rowClass:'normal bold' },
    { key:'operasional',  label:'OPERASIONAL',           type:'rp-neg',  rowClass:'normal bold' },
    { key:'iklan',        label:'IKLAN (real cost)', type:'rp-neg',  rowClass:'normal bold' },
    { key:'_h1',          label:'RASIO ADMIN & LAYANAN', type:'header',  rowClass:'header' },
    { key:'ams',          label:'↳ Biaya Komisi AMS',    type:'rp-neg',  rowClass:'sub' },
    { key:'adminFee',     label:'↳ Biaya Administrasi',  type:'rp-neg',  rowClass:'sub' },
    { key:'layanan',      label:'↳ Biaya Layanan',       type:'rp-neg',  rowClass:'sub' },
    { key:'proses',       label:'↳ Biaya Proses Pesanan',type:'rp-neg',  rowClass:'sub' },
    { key:'kampanye',     label:'↳ Biaya Kampanye',      type:'rp-neg',  rowClass:'sub' },
    { key:'_h2',          label:'METRIK KINERJA',        type:'header',  rowClass:'header' },
    { key:'aov',          label:'AOV AKTUAL',            type:'rp',      rowClass:'normal bold' },
    { key:'totalOrder',   label:'TOTAL ORDER',           type:'qty',     rowClass:'normal bold' },
    { key:'roas',         label:'ROAS AKTUAL',           type:'num',     rowClass:'normal bold' },
    { key:'acos',         label:'ACOS AKTUAL',           type:'pct',     rowClass:'normal bold' },
    { key:'_h3',          label:'PROFITABILITAS',        type:'header',  rowClass:'header' },
    { key:'gpm',          label:'Gros Profit Margin',    type:'pct',     rowClass:'normal' },
    { key:'npm',          label:'Net Profit Margin',     type:'pct',     rowClass:'normal' },
    { key:'laba',         label:'LABA / RUGI',           type:'laba',    rowClass:'laba' },
];

let rekapData = [];

async function loadRekap() {
    // Coba dari Supabase dulu
    if (typeof rekapGetAll === 'function' && typeof getAktifTokoId === 'function' && getAktifTokoId()) {
        try {
            const rows = await rekapGetAll();
            if (rows && rows.length > 0) {
                rekapData = rows.map(r => ({
                    bulan: r.periode,
                    data: {
                        pendapatan: r.total_pendapatan, penghasilan: r.total_penghasilan,
                        hpp: r.hpp, operasional: r.operasional, iklan: r.iklan,
                        ams: r.admin_ams, adminFee: r.admin_fee, layanan: r.admin_layanan,
                        proses: r.admin_proses, kampanye: r.isi_saldo,
                        aov: r.aov, totalOrder: r.total_order, roas: r.roas,
                        acos: r.roas > 0 ? (100/r.roas).toFixed(2) : 0,
                        gpm: r.gpm, npm: r.npm, laba: r.laba_rugi
                    }
                }));
                renderRekapTable();
                return;
            }
        } catch(e) { console.warn('Supabase rekap load error:', e); }
    }
    // Fallback ke localStorage
    const tokoKey = 'rekapTahunan_' + (typeof getAktifTokoNama === 'function' ? getAktifTokoNama() : 'default');
    const saved = localStorage.getItem(tokoKey) || localStorage.getItem('rekapTahunan');
    if (saved) rekapData = JSON.parse(saved);
    renderRekapTable();
}

async function saveRekap() {
    const tokoKey = 'rekapTahunan_' + (typeof getAktifTokoNama === 'function' ? getAktifTokoNama() : 'default');
    localStorage.setItem(tokoKey, JSON.stringify(rekapData));
    // Sync ke Supabase
    if (typeof rekapSave === 'function' && typeof getAktifTokoId === 'function' && getAktifTokoId()) {
        for (const b of rekapData) {
            try {
                const d = b.data;
                await rekapSave(b.bulan, {
                    total_pendapatan: parseInt(d.pendapatan)||0,
                    total_penghasilan: parseInt(d.penghasilan)||0,
                    hpp: parseInt(d.hpp)||0, operasional: parseInt(d.operasional)||0,
                    iklan: parseInt(d.iklan)||0, laba_rugi: parseInt(d.laba)||0,
                    npm: parseFloat(d.npm)||0, gpm: parseFloat(d.gpm)||0,
                    roas: parseFloat(d.roas)||0, total_order: parseInt(d.totalOrder)||0,
                    aov: parseInt(d.aov)||0, admin_ams: parseInt(d.ams)||0,
                    admin_fee: parseInt(d.adminFee)||0, admin_layanan: parseInt(d.layanan)||0,
                    admin_proses: parseInt(d.proses)||0, isi_saldo: parseInt(d.kampanye)||0
                });
            } catch(e) { console.warn('Supabase rekap save error:', e); }
        }
    }
}

function tambahBulanRekap() {
    const bulanNames = ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
    const nextIdx = rekapData.length % 12;
    const nama = prompt('Nama bulan:', bulanNames[nextIdx] + ' ' + (document.getElementById('rekap_tahun').value||new Date().getFullYear()));
    if (!nama) return;
    const emptyData = {};
    REKAP_METRIK.forEach(m => { if (m.type !== 'header') emptyData[m.key] = ''; });
    rekapData.push({ bulan: nama, data: emptyData });
    saveRekap();
    renderRekapTable();
}

function hapusBulanRekap(idx) {
    if (!confirm('Hapus bulan "' + rekapData[idx].bulan + '"?')) return;
    rekapData.splice(idx, 1);
    saveRekap();
    renderRekapTable();
}

function resetRekap() {
    if (!confirm('Reset semua data rekap tahunan?')) return;
    rekapData = [];
    const tokoKey = 'rekapTahunan_' + (typeof getAktifTokoNama === 'function' ? getAktifTokoNama() : 'default');
    localStorage.removeItem(tokoKey);
    localStorage.removeItem('rekapTahunan');
    renderRekapTable();
}

function updateRekapCell(bulanIdx, key, val) {
    if (!rekapData[bulanIdx]) return;
    rekapData[bulanIdx].data[key] = val;
    saveRekap();
    // Re-render laba row saja jika laba
    if (['pendapatan','penghasilan','hpp','operasional','iklan'].includes(key)) {
        const d = rekapData[bulanIdx].data;
        const pend = parseRekapNum(d.pendapatan);
        const peng = parseRekapNum(d.penghasilan);
        const hpp  = parseRekapNum(d.hpp);
        const opr  = parseRekapNum(d.operasional);
        const ikl  = parseRekapNum(d.iklan);
        const laba = peng - hpp - opr - ikl;
        const labaEl = document.getElementById(`rc_${bulanIdx}_laba`);
        if (labaEl) {
            labaEl.innerText = laba >= 0 ? formatRp(laba) : '('+formatRp(Math.abs(laba))+')';
            const row = labaEl.closest('tr');
            if (row) {
                row.className = laba >= 0 ? 'rekap-row-laba positif' : 'rekap-row-laba';
            }
        }
        if (pend > 0) {
            const npm = (laba/pend)*100;
            const npmEl = document.getElementById(`rc_${bulanIdx}_npm`);
            if (npmEl) npmEl.value = npm.toFixed(2);
            const gpm = pend>0?((pend-hpp)/pend)*100:0;
            const gpmEl = document.getElementById(`rc_${bulanIdx}_gpm`);
            if (gpmEl) gpmEl.value = gpm.toFixed(2);
        }
    }
}

function parseRekapNum(v) {
    if (!v) return 0;
    return parseFloat(String(v).replace(/[Rp.\s]/g,'').replace(',','.')) || 0;
}

function renderRekapTable() {
    const empty = document.getElementById('rekapEmpty');
    const thead = document.getElementById('rekapThead');
    const tbody = document.getElementById('rekapTbody');
    if (!thead || !tbody) return;

    if (rekapData.length === 0) {
        if (empty) empty.style.display = 'block';
        thead.innerHTML = '';
        tbody.innerHTML = '';
        return;
    }
    if (empty) empty.style.display = 'none';

    // Header
    let thHtml = '<tr><th class="rekap-th-metrik">METRIK</th>';
    rekapData.forEach((b, i) => {
        thHtml += `<th class="rekap-th-bulan"><div class="rekap-bulan-header"><span>${b.bulan}</span><button class="rekap-del-btn" onclick="hapusBulanRekap(${i})" title="Hapus">✕</button></div></th>`;
    });
    thHtml += '</tr>';
    thead.innerHTML = thHtml;

    // Body
    let tbHtml = '';
    let rowIdx = 0;
    REKAP_METRIK.forEach(m => {
        if (m.type === 'header') {
            tbHtml += `<tr class="rekap-row-header"><td class="rekap-td-label header-row" colspan="${rekapData.length+1}">${m.label}</td></tr>`;
            return;
        }
        const isEven = rowIdx % 2 === 0;
        rowIdx++;
        const rowCls = m.rowClass === 'laba' ? 'rekap-row-laba' : (isEven ? 'rekap-row-even' : 'rekap-row-odd');
        const labelCls = m.rowClass === 'sub' ? 'rekap-td-label sub' : 'rekap-td-label';

        tbHtml += `<tr class="${rowCls}"><td class="${labelCls}">${m.label}</td>`;

        rekapData.forEach((b, i) => {
            const val = b.data[m.key] || '';
            if (m.type === 'laba') {
                // Computed, read-only display
                const pend = parseRekapNum(b.data.pendapatan);
                const peng = parseRekapNum(b.data.penghasilan);
                const hpp  = parseRekapNum(b.data.hpp);
                const opr  = parseRekapNum(b.data.operasional);
                const ikl  = parseRekapNum(b.data.iklan);
                const laba = peng - hpp - opr - ikl;
                const isPos = laba >= 0;
                tbHtml += `<td class="rekap-td-input" style="text-align:center;font-weight:800;color:${isPos?'#166534':'#991b1b'};" id="rc_${i}_${m.key}">${laba!==0?(isPos?formatRp(laba):'('+formatRp(Math.abs(laba))+')'):'—'}</td>`;
            } else {
                const inputStyle = m.rowClass==='sub' ? 'color:#666;' : (m.rowClass==='normal bold'||m.rowClass==='normal'?'font-weight:700;color:#222;':'');
                tbHtml += `<td class="rekap-td-input"><input type="text" id="rc_${i}_${m.key}" value="${val}" placeholder="—" style="${inputStyle}" oninput="updateRekapCell(${i},'${m.key}',this.value)" onfocus="this.select()"></td>`;
            }
        });
        tbHtml += '</tr>';
    });
    tbody.innerHTML = tbHtml;
}

function prosesSemuaData() {
    const badge = document.getElementById('prosesStatusBadge');
    const detail = document.getElementById('prosesStatusDetail');
    const checklist = document.getElementById('prosesChecklist');

    // Cek status masing-masing
    const hasIncome  = rkData.income !== null;
    const hasOrder1  = rkData.order1 !== null;
    const hasOrder2  = rkData.order2 !== null;
    const hasAds     = rkData.ads    !== null;

    // Update checklist
    checklist.style.display = 'grid';
    const setChk = (id, ok, msg) => {
        const el = document.getElementById(id);
        if (!el) return;
        el.style.background = ok ? '#f0fdf4' : '#fff5f5';
        el.style.color       = ok ? '#166534' : '#991b1b';
        el.innerText         = (ok ? '✅ ' : '❌ ') + msg;
    };
    setChk('chk_income', hasIncome,  hasIncome  ? 'Income — OK' : 'Income belum diupload');
    setChk('chk_order1', hasOrder1,  hasOrder1  ? `Pesanan Ini — ${rkData.order1.totalOrder} order` : 'Pesanan bulan ini belum diupload');
    setChk('chk_order2', hasOrder2,  hasOrder2  ? `Pesanan Lalu — ${rkData.order2.totalOrder} order` : 'Pesanan bulan lalu belum diupload');
    setChk('chk_ads',    hasAds,     hasAds     ? `Ads — ${formatRp(rkData.ads.totalAds)} real cost` : 'Biaya iklan belum diisi');

    // Trigger kalkulasi
    updateRasioDashboard();

    // Tentukan status keseluruhan
    const allOk = hasIncome && hasOrder1 && hasOrder2 && hasAds;
    const someOk = hasIncome || hasOrder1;

    badge.style.display = 'block';
    if (allOk) {
        badge.style.background = '#f0fdf4';
        badge.style.color = '#166534';
        badge.style.border = '1px solid #86efac';
        badge.innerText = '✅ Semua data berhasil diproses — cek Overview untuk hasil analisis';
        detail.innerText = '';
    } else if (someOk) {
        const missing = [];
        if (!hasIncome) missing.push('Income');
        if (!hasOrder1) missing.push('Pesanan Bulan Ini');
        if (!hasOrder2) missing.push('Pesanan Bulan Lalu');
        if (!hasAds) missing.push('Biaya Iklan');
        badge.style.background = '#fef9c3';
        badge.style.color = '#854d0e';
        badge.style.border = '1px solid #fde68a';
        badge.innerText = '⚠️ Data sebagian — hasil analisis mungkin tidak lengkap';
        detail.innerText = 'Belum ada: ' + missing.join(', ');
    } else {
        badge.style.background = '#fff0f0';
        badge.style.color = '#991b1b';
        badge.style.border = '1px solid #fca5a5';
        badge.innerText = '❌ Belum ada data yang diupload';
        detail.innerText = 'Upload minimal file Income & Pesanan Bulan Ini';
    }
}

function switchTabByName(tabName) {
    const btns = document.querySelectorAll('.tab-btn');
    let targetBtn = null;
    btns.forEach(b => {
        const fn = b.getAttribute('onclick') || '';
        if (fn.includes("'" + tabName + "'")) targetBtn = b;
    });
    if (targetBtn) switchTab(tabName, targetBtn);
}


// ══════════════════════════════════════════════════════
//  ANALISIS PRODUK — parentskudetail parser & renderer
// ══════════════════════════════════════════════════════

// ── Utilities ─────────────────────────────────────────
function _skuParseAngka(val) {
    if (val === null || val === undefined || val === '-' || val === '') return 0;
    if (typeof val === 'number') return val;
    const s = String(val).replace(/\./g, '').replace(',', '.').replace(/[^0-9.]/g, '');
    return parseFloat(s) || 0;
}
function _skuParsePct(val) {
    if (!val || val === '-') return 0;
    const s = String(val).replace(',', '.').replace('%', '').trim();
    return parseFloat(s) || 0;
}
function _skuFmt(n) {
    if (!n || n === 0) return '—';
    if (n >= 1000000000) return 'Rp ' + (n/1000000000).toFixed(1) + 'M';
    if (n >= 1000000) return 'Rp ' + (n/1000000).toFixed(1) + 'jt';
    if (n >= 1000) return 'Rp ' + (n/1000).toFixed(0) + 'rb';
    return 'Rp ' + n.toLocaleString('id-ID');
}
function _skuShortName(name) {
    if (!name) return '—';
    const s = String(name);
    return s.length > 38 ? s.substring(0, 38) + '…' : s;
}

// Konversi produkArr dari rkData.performa ke format skuList
function _fromRkData(produkArr) {
    return produkArr.map(d => ({
        sku: d.skuInduk || '—',
        nama: d.namaProduk || '—',
        omset: d.penjualan || 0,
        pesanan: d.orderDibuat || 0,
        views: d.kunjungan || 0,
        keranjang: d.dimasukkan || 0,
        suka: 0,
        repeatPct: d.repeatPct || 0,
        konversiPct: d.cvrDibuat || 0,
        orderSiap: d.orderSiap || 0,
        cvrSiap: d.cvrSiap || 0
    }));
}

// Load SheetJS hanya jika perlu upload manual
function _loadSheetJS(cb) {
    if (window.XLSX) { cb(); return; }
    const s = document.createElement('script');
    s.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
    s.onload = cb;
    document.head.appendChild(s);
}

// ============================================================
// ANALISIS BIAYA LAYANAN — ENGINE v3 (pakai rkData dari Upload & Data)
// ============================================================
let biayaRawData = [];

function resetBiayaData() {
    biayaRawData = [];
    document.getElementById('biayaDashboard').style.display = 'none';
    document.getElementById('biayaUploadHint').style.display = '';
}

function syncBiayaStatusBar() {
    const hasIncomeRows    = rkData._incomeRawRows && rkData._incomeRawRows.length > 0;
    const hasIncomeSummary = rkData.income !== null && !hasIncomeRows; // format Ringkasan — ada income tapi tanpa baris per pesanan
    const hasOrder1 = rkData.order1 && rkData._orderRawMap && Object.keys(rkData._orderRawMap).length > 0;
    const hasOrder2 = !!rkData.order2;

    const el = (id, ok, warn, label) => {
        const elNode = document.getElementById(id);
        if (!elNode) return;
        elNode.style.background = ok ? '#f0fdf4' : warn ? '#fffbeb' : '#f5f5f5';
        elNode.style.color      = ok ? '#166534' : warn ? '#92400e' : '#aaa';
        elNode.textContent      = ok ? '✅ ' + label : warn ? '⚠️ ' + label : '⬜ ' + label;
    };

    if (hasIncomeRows) {
        el('biayaStIncome', true, false, `Income (${rkData._incomeRawRows.length} pesanan)`);
    } else if (hasIncomeSummary) {
        el('biayaStIncome', false, true, `Income (format Ringkasan — butuh file per transaksi)`);
    } else {
        el('biayaStIncome', false, false, `Income`);
    }
    el('biayaStOrder1', hasOrder1, false, `Pesanan Bulan Ini`);
    el('biayaStOrder2', hasOrder2, false, `Pesanan Bulan Lalu`);
}

function prosesBiayaFromRkData() {
    syncBiayaStatusBar();
    const rows = rkData._incomeRawRows || [];
    if (!rows.length) {
        // Cek apakah income ada tapi format Ringkasan
        if (rkData.income !== null) {
            alert('File Income yang diupload adalah format Ringkasan (Summary).\n\nUntuk Analisis Biaya Layanan, diperlukan file Income format per-transaksi (baris per pesanan).\n\nCara mendapatkannya:\n• Buka Seller Center Shopee\n• Keuangan → Laporan Keuangan\n• Pilih "Data Income" atau unduh laporan transaksi harian\n• Pastikan file berisi kolom: No. Pesanan, Total Penghasilan, dll.');
        } else {
            alert('Belum ada data Income. Silakan upload file Income di tab Upload & Data terlebih dahulu.');
        }
        switchTabByName('rasio_upload');
        return;
    }

    const BIAYA_NAMES = ['Biaya Komisi AMS','Biaya Administrasi','Biaya Layanan','Biaya Proses Pesanan','Biaya Kampanye','Biaya Program Hemat Biaya Kirim','Biaya Transaksi','Premi'];
    const ordMap = rkData._orderRawMap || {};

    biayaRawData = rows.map(r => {
        const vchrShopeeOrder = (ordMap[r.noPesanan]||{}).vchrShopee || 0;
        const totalVchrProgram = vchrShopeeOrder + Math.abs(r.voucherSeller||0);
        const ditanggungShopee = vchrShopeeOrder + Math.abs(r.gofs||0);
        const ditanggungSeller = Math.abs(r.voucherSeller||0);
        return {
            ...r,
            vchrShopeeOrder,
            totalVchrProgram,
            ditanggungShopee,
            ditanggungSeller,
            rasioBiaya:  r.omset > 0 ? Math.abs(r.totalBiaya||0)/r.omset*100 : 0,
            rasioPotongan: r.omset > 0 ? Math.abs(r.totalPotonganSeller||0)/r.omset*100 : 0,
        };
    });

    const n = biayaRawData.length;
    document.getElementById('biayaFileInfo').textContent =
        `${n} pesanan Income | Order map: ${Object.keys(ordMap).length} entri`;
    renderBiayaDashboard();
}

function fmtRp(v) {
    return new Intl.NumberFormat('id-ID',{style:'currency',currency:'IDR',minimumFractionDigits:0,maximumFractionDigits:0}).format(Math.abs(Math.floor(v||0)));
}

function renderBiayaDashboard() {
    document.getElementById('biayaUploadHint').style.display = 'none';
    document.getElementById('biayaDashboard').style.display = '';

    const d = biayaRawData;
    const n = d.length;
    // totalOmset = pendapatan NETT dari Income (selesai & cair, tanpa batal)
    // Konsisten dengan Rasio Keuangan — hanya data yang real masuk
    const totalOmset    = (rkData.income && rkData.income.totalPendapatan > 0)
                            ? rkData.income.totalPendapatan
                            : d.reduce((s,r)=>s+r.omset,0);
    const totalIncome   = d.reduce((s,r)=>s+r.income,0);
    const totalBiaya    = d.reduce((s,r)=>s+(r.totalBiaya||0),0);
    const totalPotongan = d.reduce((s,r)=>s+(r.totalPotonganSeller||0),0);
    const totalGofs     = d.reduce((s,r)=>s+(r.gofs||0),0);
    const totalVchrShopee = d.reduce((s,r)=>s+(r.vchrShopeeOrder||0),0);

    // ── SUMMARY CARDS ──
    document.getElementById('biayaSummaryCards').innerHTML = [
        { lbl:'Total Pendapatan (Omset)', val:fmtRp(totalOmset),           sub:`${n} pesanan selesai & cair`,       color:'#1a1a2e', bg:'#f8f9ff' },
        { lbl:'Total Income Cair',        val:fmtRp(totalIncome),           sub:`Rata-rata ${fmtRp(totalIncome/n)}/pesanan`,  color:'#166534', bg:'#f0fdf4' },
        { lbl:'Total Biaya Layanan',      val:fmtRp(Math.abs(totalBiaya)), sub:`Rata-rata ${fmtRp(Math.abs(totalBiaya)/n)}/pesanan`, color:'#991b1b', bg:'#fff1f2' },
        { lbl:'Total Potongan Seller',    val:fmtRp(Math.abs(totalPotongan)),sub:'Diskon produk + voucher seller',   color:'#92400e', bg:'#fffbeb' },
    ].map(c=>`<div style="background:${c.bg};border-radius:10px;padding:12px 14px;border:1px solid rgba(0,0,0,0.05);">
        <div style="font-size:0.63em;font-weight:700;color:#888;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:4px;">${c.lbl}</div>
        <div style="font-size:1.05em;font-weight:800;color:${c.color};">${c.val}</div>
        <div style="font-size:0.62em;color:#aaa;margin-top:2px;">${c.sub}</div>
    </div>`).join('');

    // ── RASIO LEVERAGE ──
    const r1=(Math.abs(totalBiaya)/totalOmset*100), r2=(Math.abs(totalBiaya)/totalIncome*100);
    const r3=(Math.abs(totalPotongan)/totalOmset*100), r4=(totalGofs/totalOmset*100);
    document.getElementById('biayaRasioCards').innerHTML = [
        { lbl:'Biaya Layanan / Omset',   val:r1.toFixed(2)+'%', desc:'Beban platform dari harga asli',  warn:r1>13 },
        { lbl:'Biaya Layanan / Income',  val:r2.toFixed(2)+'%', desc:'Beban dari uang yang cair',        warn:r2>35 },
        { lbl:'Potongan Seller / Omset', val:r3.toFixed(2)+'%', desc:'Total keluar dari kantong seller', warn:r3>45 },
        { lbl:'Subsidi GOFS / Omset',    val:r4.toFixed(2)+'%', desc:`Subsidi ongkir Shopee: ${fmtRp(totalGofs)}`, warn:false },
    ].map(c=>`<div style="background:${c.warn?'rgba(239,68,68,0.12)':'rgba(255,255,255,0.07)'};border-radius:10px;padding:12px 14px;border:1px solid ${c.warn?'rgba(239,68,68,0.3)':'rgba(255,255,255,0.06)'};">
        <div style="font-size:0.6em;font-weight:700;color:rgba(255,255,255,0.4);text-transform:uppercase;letter-spacing:0.5px;margin-bottom:4px;">${c.lbl}</div>
        <div style="font-size:1.25em;font-weight:800;color:${c.warn?'#fca5a5':'#fff'};">${c.val}</div>
        <div style="font-size:0.6em;color:rgba(255,255,255,0.32);margin-top:3px;">${c.desc}</div>
        ${c.warn?'<div style="font-size:0.6em;color:#fca5a5;margin-top:4px;font-weight:700;">⚠️ Di atas batas wajar</div>':''}
    </div>`).join('');

    // ── BREAKDOWN BIAYA ──
    const BIAYA_COLS=['Biaya Komisi AMS','Biaya Administrasi','Biaya Layanan','Biaya Proses Pesanan','Biaya Kampanye','Biaya Program Hemat Biaya Kirim','Biaya Transaksi','Premi'];
    const totals={}; BIAYA_COLS.forEach(c=>{totals[c]=d.reduce((s,r)=>s+((r.biayaDetail||{})[c]||0),0);});
    const aktif=BIAYA_COLS.filter(c=>totals[c]!==0);
    const sumAbs=aktif.reduce((s,c)=>s+Math.abs(totals[c]),0);
    document.getElementById('biayaBreakdown').innerHTML=aktif.map(c=>{
        const abs=Math.abs(totals[c]);
        const pct=sumAbs>0?(abs/sumAbs*100):0;
        const pctO=totalOmset>0?(abs/totalOmset*100):0;
        return `<div style="margin-bottom:10px;">
            <div style="display:flex;justify-content:space-between;margin-bottom:3px;">
                <span style="font-size:0.71em;font-weight:600;color:#444;">${c}</span>
                <span style="font-size:0.71em;font-weight:800;color:#ee4d2d;">${fmtRp(abs)}</span>
            </div>
            <div style="background:#f5f5f5;border-radius:4px;height:5px;overflow:hidden;">
                <div style="background:#ee4d2d;height:100%;width:${pct.toFixed(1)}%;border-radius:4px;"></div>
            </div>
            <div style="display:flex;justify-content:space-between;margin-top:2px;">
                <span style="font-size:0.61em;color:#bbb;">${pct.toFixed(1)}% dari total biaya</span>
                <span style="font-size:0.61em;color:#bbb;">${pctO.toFixed(2)}% dari omset</span>
            </div>
        </div>`;
    }).join('')+`<div style="margin-top:10px;padding-top:8px;border-top:2px solid #fee2e2;display:flex;justify-content:space-between;">
        <span style="font-size:0.72em;font-weight:800;color:#1a1a2e;">TOTAL</span>
        <span style="font-size:0.72em;font-weight:800;color:#ee4d2d;">${fmtRp(sumAbs)} (${(sumAbs/totalOmset*100).toFixed(2)}% omset)</span>
    </div>`;

    // ── TABEL RINGKAS PER PROGRAM ──
    const PROGS=['Vchr Extra','Video Extra','Live Extra','GOFS Extra','Voucher Penjual','Tanpa Program'];
    const PC={'Vchr Extra':'#3730a3','Video Extra':'#0d9488','Live Extra':'#dc2626','GOFS Extra':'#0369a1','Voucher Penjual':'#92400e','Tanpa Program':'#6b7280'};
    let ph=`<table style="width:100%;border-collapse:collapse;font-size:0.7em;">
        <thead><tr style="background:#f8f9ff;">
            <th style="padding:5px 7px;text-align:left;color:#555;border-bottom:2px solid #e0e7ff;">Program</th>
            <th style="padding:5px 7px;text-align:center;color:#555;border-bottom:2px solid #e0e7ff;">Pesanan</th>
            <th style="padding:5px 7px;text-align:right;color:#555;border-bottom:2px solid #e0e7ff;">Omset</th>
            <th style="padding:5px 7px;text-align:right;color:#555;border-bottom:2px solid #e0e7ff;">Biaya Layanan</th>
            <th style="padding:5px 7px;text-align:right;color:#555;border-bottom:2px solid #e0e7ff;">% Omset</th>
            <th style="padding:5px 7px;text-align:right;color:#555;border-bottom:2px solid #e0e7ff;">Subsidi Shopee</th>
            <th style="padding:5px 7px;text-align:right;color:#555;border-bottom:2px solid #e0e7ff;">Income Cair</th>
        </tr></thead><tbody>`;
    PROGS.forEach(prog=>{
        const rows=d.filter(r=>r.program===prog); if(!rows.length)return;
        const pO=rows.reduce((s,r)=>s+r.omset,0), pB=rows.reduce((s,r)=>s+(r.totalBiaya||0),0);
        const pG=rows.reduce((s,r)=>s+(r.gofs||0),0), pVS=rows.reduce((s,r)=>s+(r.vchrShopeeOrder||0),0);
        const pI=rows.reduce((s,r)=>s+r.income,0);
        const rs=pO>0?(Math.abs(pB)/pO*100).toFixed(2):'0.00';
        ph+=`<tr style="border-bottom:1px solid #f0f0f0;">
            <td style="padding:6px 7px;font-weight:700;color:${PC[prog]||'#555'};">${prog}<div style="font-size:0.82em;color:#bbb;font-weight:500;">${rows.length} pesanan (${(rows.length/n*100).toFixed(0)}%)</div></td>
            <td style="padding:6px 7px;text-align:center;"><span style="background:${PC[prog]||'#555'}22;color:${PC[prog]||'#555'};border-radius:20px;padding:1px 9px;font-weight:800;">${rows.length}</span></td>
            <td style="padding:6px 7px;text-align:right;font-weight:600;">${fmtRp(pO)}</td>
            <td style="padding:6px 7px;text-align:right;color:#ee4d2d;font-weight:700;">${fmtRp(Math.abs(pB))}</td>
            <td style="padding:6px 7px;text-align:right;font-weight:800;color:${parseFloat(rs)>13?'#dc2626':'#1a1a2e'};">${rs}%</td>
            <td style="padding:6px 7px;text-align:right;color:#0d9488;font-weight:600;">${fmtRp(pG+pVS)}</td>
            <td style="padding:6px 7px;text-align:right;color:#166534;font-weight:700;">${fmtRp(pI)}</td>
        </tr>`;
    });
    ph+=`<tr style="background:#f8f8f8;font-weight:800;border-top:2px solid #e0e7ff;">
        <td style="padding:6px 7px;">TOTAL</td><td style="padding:6px 7px;text-align:center;">${n}</td>
        <td style="padding:6px 7px;text-align:right;">${fmtRp(totalOmset)}</td>
        <td style="padding:6px 7px;text-align:right;color:#ee4d2d;">${fmtRp(Math.abs(totalBiaya))}</td>
        <td style="padding:6px 7px;text-align:right;">${(Math.abs(totalBiaya)/totalOmset*100).toFixed(2)}%</td>
        <td style="padding:6px 7px;text-align:right;color:#0d9488;">${fmtRp(totalGofs+totalVchrShopee)}</td>
        <td style="padding:6px 7px;text-align:right;color:#166534;">${fmtRp(totalIncome)}</td>
    </tr></tbody></table>`;
    document.getElementById('biayaProgramTable').innerHTML=ph;

    // ── KARTU ANALISIS PER PROGRAM — Grid 2 kolom, simple + informatif + eksekusi ──
    const PROGS_SHOW = PROGS.filter(p => d.filter(r=>r.program===p).length > 0 && p !== 'Tanpa Program');
    // Simpan data program untuk export PDF
    window._biayaProgData = [];
    document.getElementById('biayaProgramCards').innerHTML = PROGS_SHOW.map(prog => {
        const rows = d.filter(r => r.program === prog);
        const nPakai = rows.length;
        const pctPakai = (nPakai / n * 100).toFixed(1);
        const pOmset = rows.reduce((s,r) => s + r.omset, 0);
        const totalVchr = rows.reduce((s,r) => s + (r.totalVchrProgram||0), 0);
        const ditShopee = rows.reduce((s,r) => s + (r.ditanggungShopee||0), 0);
        const ditSeller = rows.reduce((s,r) => s + (r.ditanggungSeller||0), 0);
        const rataVchr = nPakai > 0 ? totalVchr / nPakai : 0;
        const pctSh = totalVchr > 0 ? (ditShopee / totalVchr * 100) : 0;
        const pctSl = totalVchr > 0 ? (ditSeller / totalVchr * 100) : 0;
        const rasio = totalOmset > 0 ? (ditSeller / totalOmset * 100) : 0;
        const leverRp = ditSeller > 0 ? Math.round(ditShopee / ditSeller * 1000) : 0;
        const col = PC[prog] || '#3730a3';
        const warn = rasio > 5;
        const simOmsetTanpa = totalOmset - pOmset;
        const simSelisih = pOmset;
        // Simpan untuk PDF
        window._biayaProgData.push({prog,n,nPakai,pctPakai,totalOmset,pOmset,totalVchr,ditShopee,ditSeller,rataVchr,pctSh,pctSl,rasio,leverRp,warn});
        const progId = prog.replace(/\s+/g,'_').toLowerCase();
        return `<div style="background:#fff;border:1px solid #e8e8e8;border-radius:14px;padding:18px;border-top:4px solid ${col};box-shadow:0 1px 4px rgba(0,0,0,0.05);">

            <!-- HEADER -->
            <div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:14px;">
                <div>
                    <div style="font-size:0.85em;font-weight:800;color:${col};">${prog}</div>
                    <div style="font-size:0.65em;color:#aaa;margin-top:2px;">${nPakai} dari ${n} pesanan · ${pctPakai}%</div>
                </div>
                <div style="background:${warn?'#fff1f2':'#f0fdf4'};border:1.5px solid ${warn?'#fca5a5':'#86efac'};border-radius:10px;padding:6px 14px;text-align:center;min-width:80px;">
                    <div style="font-size:0.58em;color:${warn?'#be123c':'#16a34a'};font-weight:800;text-transform:uppercase;letter-spacing:0.5px;">Rasio Beban</div>
                    <div style="font-size:1.4em;font-weight:800;color:${warn?'#9f1239':'#166534'};line-height:1.1;">${rasio.toFixed(1)}%</div>
                    <div style="font-size:0.56em;color:#bbb;">Biaya / Pendapatan</div>
                </div>
            </div>

            <!-- METRICS GRID — 3 kolom x 2 baris -->
            <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px;margin-bottom:12px;">
                <div style="background:#f8f9ff;border-radius:8px;padding:9px 10px;text-align:center;">
                    <div style="font-size:0.58em;color:#888;font-weight:700;text-transform:uppercase;margin-bottom:2px;">Total Pendapatan</div>
                    <div style="font-size:0.88em;font-weight:800;color:#1a1a2e;">${fmtRp(totalOmset)}</div>
                    <div style="font-size:0.58em;color:#bbb;">${n} pesanan</div>
                </div>
                <div style="background:#f8f9ff;border-radius:8px;padding:9px 10px;text-align:center;">
                    <div style="font-size:0.58em;color:#888;font-weight:700;text-transform:uppercase;margin-bottom:2px;">Total Voucher</div>
                    <div style="font-size:0.88em;font-weight:800;color:#1a1a2e;">${fmtRp(totalVchr)}</div>
                    <div style="font-size:0.58em;color:#bbb;">${totalOmset>0?(totalVchr/totalOmset*100).toFixed(1):'0'}% dari omset</div>
                </div>
                <div style="background:#f0fdf4;border-radius:8px;padding:9px 10px;text-align:center;">
                    <div style="font-size:0.58em;color:#16a34a;font-weight:700;text-transform:uppercase;margin-bottom:2px;">Rata-rata Vchr</div>
                    <div style="font-size:0.88em;font-weight:800;color:#166534;">${fmtRp(rataVchr)}</div>
                    <div style="font-size:0.58em;color:#bbb;">${nPakai} pesanan pakai</div>
                </div>
                <div style="background:#fff1f2;border-radius:8px;padding:9px 10px;text-align:center;">
                    <div style="font-size:0.58em;color:#be123c;font-weight:700;text-transform:uppercase;margin-bottom:2px;">Potongan Seller</div>
                    <div style="font-size:0.88em;font-weight:800;color:#9f1239;">${fmtRp(ditSeller)}</div>
                    <div style="font-size:0.58em;color:#bbb;">${pctSl.toFixed(1)}% dari total vchr</div>
                </div>
                <div style="background:#f0f9ff;border-radius:8px;padding:9px 10px;text-align:center;">
                    <div style="font-size:0.58em;color:#0369a1;font-weight:700;text-transform:uppercase;margin-bottom:2px;">Ditanggung Shopee</div>
                    <div style="font-size:0.88em;font-weight:800;color:#075985;">${fmtRp(ditShopee)}</div>
                    <div style="font-size:0.58em;color:#bbb;">${pctSh.toFixed(1)}% dari total vchr</div>
                </div>
                <div style="background:#fefce8;border-radius:8px;padding:9px 10px;text-align:center;">
                    <div style="font-size:0.58em;color:#a16207;font-weight:700;text-transform:uppercase;margin-bottom:2px;">Leverage Shopee</div>
                    <div style="font-size:0.88em;font-weight:800;color:#92400e;">${fmtRp(leverRp)}</div>
                    <div style="font-size:0.58em;color:#bbb;">per Rp1.000 seller keluar</div>
                </div>
            </div>

            <!-- RINGKASAN ANALISIS -->
            <div style="background:${warn?'#fff9f9':'#f8faff'};border-radius:8px;padding:11px 13px;margin-bottom:12px;border-left:3px solid ${warn?'#fca5a5':col};">
                <div style="font-size:0.65em;font-weight:800;color:${col};text-transform:uppercase;letter-spacing:0.4px;margin-bottom:5px;">📝 Ringkasan Analisis</div>
                <div style="font-size:0.72em;line-height:1.75;color:#555;">
                    Dari <b>${n} pesanan</b>, sebanyak <b>${nPakai} pesanan (${pctPakai}%)</b> menggunakan <b>${prog}</b>.
                    Rata-rata pembeli mendapat potongan <b>${fmtRp(rataVchr)}</b> per pesanan.
                    Dari total voucher, <b>${pctSh.toFixed(1)}%</b> disubsidi Shopee &amp; <b>${pctSl.toFixed(1)}%</b> ditanggung seller.
                    Untuk setiap <b>Rp 1.000</b> yang kamu keluarkan, Shopee memberikan subsidi <b>${fmtRp(leverRp)}</b>.
                </div>
                <div style="font-size:0.68em;font-weight:700;margin-top:6px;color:${warn?'#dc2626':'#0d9488'};">
                    ${warn?'⚠️ Rasio beban di atas 5% — pertimbangkan evaluasi ulang program ini.':'✅ Rasio beban masih wajar — program memberikan leverage subsidi yang baik.'}
                </div>
            </div>

            <!-- SIMULASI -->
            <div style="background:#f8f8f8;border-radius:8px;padding:11px 13px;margin-bottom:12px;">
                <div style="font-size:0.65em;font-weight:800;color:#555;text-transform:uppercase;letter-spacing:0.4px;margin-bottom:8px;">🔮 Simulasi — Jika Program Dimatikan</div>
                <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;">
                    <div style="background:#fff;border-radius:7px;padding:8px 10px;border:1px solid #eee;">
                        <div style="font-size:0.6em;color:#888;text-transform:uppercase;font-weight:700;margin-bottom:2px;">Omset yang Mungkin Hilang</div>
                        <div style="font-size:0.92em;font-weight:800;color:#dc2626;">- ${fmtRp(simSelisih)}</div>
                        <div style="font-size:0.58em;color:#bbb;">estimasi dari ${nPakai} pesanan program ini</div>
                    </div>
                    <div style="background:#fff;border-radius:7px;padding:8px 10px;border:1px solid #eee;">
                        <div style="font-size:0.6em;color:#888;text-transform:uppercase;font-weight:700;margin-bottom:2px;">Biaya yang Dihemat</div>
                        <div style="font-size:0.92em;font-weight:800;color:#166534;">+ ${fmtRp(ditSeller)}</div>
                        <div style="font-size:0.58em;color:#bbb;">potongan seller tidak lagi keluar</div>
                    </div>
                </div>
            </div>

            <!-- TOMBOL EKSEKUSI -->
            <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;">
                <button onclick="eksekusiProgram('lanjut','${prog}')" style="background:#166534;color:#fff;border:none;padding:9px;border-radius:8px;font-size:0.72em;font-weight:800;cursor:pointer;">✅ Lanjutkan Program</button>
                <button onclick="eksekusiProgram('berhenti','${prog}')" style="background:#fff;color:#dc2626;border:1.5px solid #fca5a5;padding:9px;border-radius:8px;font-size:0.72em;font-weight:800;cursor:pointer;">⛔ Berhenti dari Program</button>
            </div>

        </div>`;
    }).join('');

    renderBiayaDetailTable();
}

function renderBiayaDetailTable() {
    const filter=document.getElementById('biayaFilterProgram').value;
    const d=filter==='all'?biayaRawData:biayaRawData.filter(r=>r.program===filter);
    const PC={'Vchr Extra':'#3730a3','Video Extra':'#0d9488','Live Extra':'#dc2626','GOFS Extra':'#0369a1','Voucher Penjual':'#92400e','Tanpa Program':'#6b7280'};
    let html=`<table style="width:100%;border-collapse:collapse;font-size:0.7em;min-width:800px;">
        <thead><tr style="background:#f8f8f8;">
            <th style="padding:6px 8px;text-align:left;font-weight:800;color:#555;border-bottom:2px solid #eee;white-space:nowrap;position:sticky;top:0;background:#f8f8f8;">No. Pesanan</th>
            <th style="padding:6px 8px;text-align:left;font-weight:800;color:#555;border-bottom:2px solid #eee;position:sticky;top:0;background:#f8f8f8;">Program</th>
            <th style="padding:6px 8px;text-align:right;font-weight:800;color:#555;border-bottom:2px solid #eee;white-space:nowrap;position:sticky;top:0;background:#f8f8f8;">Omset</th>
            <th style="padding:6px 8px;text-align:right;font-weight:800;color:#555;border-bottom:2px solid #eee;white-space:nowrap;position:sticky;top:0;background:#f8f8f8;">Ptg. Seller</th>
            <th style="padding:6px 8px;text-align:right;font-weight:800;color:#555;border-bottom:2px solid #eee;white-space:nowrap;position:sticky;top:0;background:#f8f8f8;">Biaya Layanan</th>
            <th style="padding:6px 8px;text-align:right;font-weight:800;color:#555;border-bottom:2px solid #eee;white-space:nowrap;position:sticky;top:0;background:#f8f8f8;">Subsidi Shopee</th>
            <th style="padding:6px 8px;text-align:right;font-weight:800;color:#555;border-bottom:2px solid #eee;white-space:nowrap;position:sticky;top:0;background:#f8f8f8;">Income Cair</th>
            <th style="padding:6px 8px;text-align:center;font-weight:800;color:#555;border-bottom:2px solid #eee;white-space:nowrap;position:sticky;top:0;background:#f8f8f8;">% Beban</th>
        </tr></thead><tbody>`;
    d.forEach((r,i)=>{
        const warn=r.rasioBiaya>13;
        html+=`<tr style="border-bottom:1px solid #f5f5f5;background:${i%2===0?'#fff':'#fafafa'};">
            <td style="padding:5px 8px;font-weight:700;color:#1a1a2e;white-space:nowrap;font-size:0.88em;">${r.noPesanan}</td>
            <td style="padding:5px 8px;"><span style="background:${PC[r.program]||'#555'}18;color:${PC[r.program]||'#555'};border-radius:4px;padding:1px 7px;font-weight:700;">${r.program}</span></td>
            <td style="padding:5px 8px;text-align:right;">${fmtRp(r.omset)}</td>
            <td style="padding:5px 8px;text-align:right;color:#92400e;">${fmtRp(Math.abs(r.totalPotonganSeller||0))}</td>
            <td style="padding:5px 8px;text-align:right;color:#ee4d2d;font-weight:700;">${fmtRp(Math.abs(r.totalBiaya||0))}</td>
            <td style="padding:5px 8px;text-align:right;color:#0d9488;">${fmtRp(r.ditanggungShopee||0)}</td>
            <td style="padding:5px 8px;text-align:right;color:#166534;font-weight:700;">${fmtRp(r.income)}</td>
            <td style="padding:5px 8px;text-align:center;"><span style="background:${warn?'#fee2e2':'#f0fdf4'};color:${warn?'#dc2626':'#16a34a'};border-radius:4px;padding:2px 8px;font-weight:800;">${r.rasioBiaya.toFixed(2)}%</span></td>
        </tr>`;
    });
    html+='</tbody></table>';
    document.getElementById('biayaDetailTableWrap').innerHTML=html;
    document.getElementById('biayaTableInfo').textContent=`Menampilkan ${d.length} pesanan${filter!=='all'?' — Filter: '+filter:''}`;
}

// ============================================================
// Dipanggil oleh switchTab saat masuk tab analisis_produk// ============================================================
// Dipanggil oleh switchTab saat masuk tab analisis_produk// ============================================================
// Dipanggil oleh switchTab saat masuk tab analisis_produk
function tryAutoLoadSkuData() {
    // rkData adalah let di scope script yang sama — langsung accessible
    if (typeof rkData !== 'undefined' && rkData && rkData.performa && rkData.performa.produkArr && rkData.performa.produkArr.length > 0) {
        const skuList = _fromRkData(rkData.performa.produkArr);
        const produkAktif = rkData.performa.produkAktif || rkData.performa.produkArr.length;
        _renderSkuDashboard(skuList, produkAktif + ' produk aktif — data dari Upload & Data', true);
    } else {
        document.getElementById('skuUploadHint').style.display = 'block';
        document.getElementById('skuDashboard').style.display = 'none';
    }
}

// Upload manual (fallback)
function prosesSkuFile(input) {
    const file = input.files[0];
    if (!file) return;
    _loadSheetJS(function() {
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const wb = XLSX.read(e.target.result, { type: 'array' });
                const sheetName = wb.SheetNames.find(n => n.toLowerCase().includes('performa')) || wb.SheetNames[0];
                const ws = wb.Sheets[sheetName];
                const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
                if (!rows || rows.length === 0) { alert('Data tidak ditemukan di sheet: ' + sheetName); return; }

                const skuMap = {};
                rows.forEach(r => {
                    const sku = r['SKU Induk'] || r['Kode Produk'] || '—';
                    const nama = r['Produk'] || '—';
                    if (!skuMap[sku]) {
                        skuMap[sku] = { sku, nama, omset:0, pesanan:0, views:0, keranjang:0, suka:0, repeatPct:0, konversiPct:0, orderSiap:0, cvrSiap:0, _rArr:[], _kArr:[] };
                    }
                    const d = skuMap[sku];
                    d.omset     += _skuParseAngka(r['Total Penjualan (Pesanan Dibuat) (IDR)']);
                    d.pesanan   += _skuParseAngka(r['Pesanan Dibuat']);
                    d.orderSiap += _skuParseAngka(r['Pesanan Siap Dikirim']);
                    d.views     += _skuParseAngka(r['Jumlah Produk Dilihat']);
                    d.keranjang += _skuParseAngka(r['Dimasukkan ke Keranjang (Produk)']);
                    d.suka      += _skuParseAngka(r['Suka']);
                    d._rArr.push(_skuParsePct(r['Tingkat Pesanan Berulang (Pesanan Dibuat)']));
                    d._kArr.push(_skuParsePct(r['Tingkat Konversi Pesanan (Pesanan Dibuat)']));
                });

                const skuList = Object.values(skuMap).map(d => {
                    d.repeatPct   = d._rArr.reduce((a,b)=>a+b,0) / (d._rArr.length || 1);
                    d.konversiPct = d._kArr.filter(v=>v>0).reduce((a,b)=>a+b,0) / (d._kArr.filter(v=>v>0).length || 1);
                    return d;
                });

                _renderSkuDashboard(skuList, file.name, false);
            } catch(err) { alert('Gagal membaca file: ' + err.message); }
        };
        reader.readAsArrayBuffer(file);
    });
}

function resetSkuData() {
    document.getElementById('skuUploadHint').style.display = 'block';
    document.getElementById('skuDashboard').style.display = 'none';
    const inp = document.getElementById('uploadSkuFile');
    if (inp) inp.value = '';
}

function _renderSkuDashboard(list, fileInfo, dariRkData) {
    document.getElementById('skuUploadHint').style.display = 'none';
    document.getElementById('skuDashboard').style.display = 'block';

    // Badge sumber data
    const badgeEl = document.getElementById('skuSumberBadge');
    if (badgeEl) {
        badgeEl.innerHTML = dariRkData
            ? '<span style="background:#f0fdf4;color:#166534;border:1px solid #86efac;padding:3px 10px;border-radius:20px;font-size:0.72em;font-weight:700;">✅ Terhubung ke Upload & Data</span>'
            : '<span style="background:#fef9c3;color:#854d0e;border:1px solid #fde68a;padding:3px 10px;border-radius:20px;font-size:0.72em;font-weight:700;">📂 Upload manual</span>';
    }
    const uploadBtn = document.getElementById('skuManualUploadBtn');
    if (uploadBtn) uploadBtn.style.display = dariRkData ? 'none' : '';

    // Omset & Pesanan: pakai dari Income (real cair, tanpa batal) jika tersedia
    // Konsisten dgn konsep: yg batal tidak masuk
    const _omsetPerforma = list.reduce((a,d)=>a+d.omset, 0);
    const totalOmset  = (rkData.income && rkData.income.totalPendapatan > 0)
                          ? rkData.income.totalPendapatan : _omsetPerforma;
    const totalPesananCair = rkData.order1?.totalOrder || 0;
    const totalOmsetStr   = "Rp " + Math.round(totalOmset).toLocaleString("id-ID");
    const totalPesanan  = list.reduce((a,d)=>a+d.pesanan, 0);
    const totalViews    = list.reduce((a,d)=>a+d.views, 0);
    const totalSku      = list.length;
    const totalSiap     = list.reduce((a,d)=>a+d.orderSiap, 0);
    const cvrAktual     = totalViews > 0 ? (totalSiap / totalViews * 100) : 0;

    // Summary cards
    document.getElementById('skuSummaryCards').innerHTML = [
        { label:'Total Omset',   val: totalOmsetStr,                                          color:'#3730a3', icon:'💰', sub: totalPesananCair > 0 ? totalPesananCair + ' pesanan cair' : '' },
        { label:'Total Pesanan',  val: (totalPesananCair > 0 ? totalPesananCair : totalPesanan).toLocaleString('id-ID'), color:'#0d9488', icon:'📦', sub: totalPesananCair > 0 ? 'selesai & cair' : 'pesanan dibuat' },
        { label:'Total Views',      val: totalViews.toLocaleString('id-ID'),            color:'#ea580c', icon:'👁' },
        { label:'CVR Toko Aktual',  val: cvrAktual.toFixed(2) + '%',                   color:'#7c3aed', icon:'📈',
          sub: totalSiap.toLocaleString('id-ID') + ' siap kirim' },
        { label:'Jumlah SKU',       val: totalSku,                                      color:'#166534', icon:'🏷️' },
    ].map(c => {
        // Buat warna tint transparan dari warna aksen untuk background kanan
        const hexToRgb = hex => {
            const r = parseInt(hex.slice(1,3),16);
            const g = parseInt(hex.slice(3,5),16);
            const b = parseInt(hex.slice(5,7),16);
            return `${r},${g},${b}`;
        };
        const rgb = hexToRgb(c.color);
        const tintBg = `rgba(${rgb},0.07)`;
        const tintBorder = `rgba(${rgb},0.2)`;
        const valColor = c.color;

        return `
        <div style="display:flex;align-items:stretch;border-radius:11px;overflow:hidden;border:1.5px solid ${tintBorder};box-shadow:0 1px 4px rgba(0,0,0,0.06);min-width:0;">
            <!-- Ikon kiri — fixed width proporsional -->
            <div style="background:${c.color};width:90px;min-width:90px;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:6px;padding:12px 6px;">
                <div style="font-size:1.6em;line-height:1;">${c.icon}</div>
                <div style="font-size:0.65em;font-weight:800;color:#fff;text-transform:uppercase;letter-spacing:0.5px;text-align:center;line-height:1.35;padding:0 4px;">${c.label}</div>
            </div>
            <!-- Nilai kanan — tint + angka besar -->
            <div style="flex:1;background:${tintBg};display:flex;flex-direction:column;align-items:center;justify-content:center;padding:10px 12px;min-width:0;text-align:center;">
                <div style="font-size:1.3em;font-weight:800;color:${valColor};line-height:1.1;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:100%;">${c.val}</div>
                ${c.sub ? `<div style="font-size:0.62em;color:#666;font-weight:600;margin-top:4px;">${c.sub}</div>` : ''}
            </div>
        </div>
    `}).join('');

    // Ongkir Seller & Pembatalan
    renderOngkirPembatalan();

    // Top 5 Omset
    const topOmset = [...list].sort((a,b)=>b.omset-a.omset).slice(0,5);
    document.getElementById('skuTopOmset').innerHTML = topOmset.map((d,i) => `
        <div style="display:flex;justify-content:space-between;align-items:center;padding:6px 0;border-bottom:1px solid #f5f5f5;font-size:0.82em;">
            <div style="display:flex;align-items:center;gap:7px;">
                <span style="background:#fee2e2;color:#ee4d2d;border-radius:5px;padding:1px 7px;font-size:0.78em;font-weight:800;">#${i+1}</span>
                <span style="color:#333;font-weight:600;">${_skuShortName(d.sku)}</span>
            </div>
            <span style="font-weight:800;color:#ee4d2d;white-space:nowrap;margin-left:8px;">${_skuFmt(d.omset)}</span>
        </div>
    `).join('');

    // Top 5 Konversi
    const avgKonv = list.reduce((a,d)=>a+d.konversiPct,0) / (list.length||1);
    const topKonv = [...list].filter(d=>d.pesanan>0||d.konversiPct>0).sort((a,b)=>b.konversiPct-a.konversiPct).slice(0,5);
    document.getElementById('skuTopKonversi').innerHTML = topKonv.length === 0
        ? '<div style="color:#bbb;font-size:0.82em;text-align:center;padding:16px;">Tidak ada data konversi.</div>'
        : topKonv.map((d,i) => `
        <div style="display:flex;justify-content:space-between;align-items:center;padding:6px 0;border-bottom:1px solid #f5f5f5;font-size:0.82em;">
            <div style="display:flex;align-items:center;gap:7px;">
                <span style="background:#ccfbf1;color:#0d9488;border-radius:5px;padding:1px 7px;font-size:0.78em;font-weight:800;">#${i+1}</span>
                <span style="color:#333;font-weight:600;">${_skuShortName(d.sku)}</span>
            </div>
            <span style="font-weight:800;color:#0d9488;white-space:nowrap;margin-left:8px;">${d.konversiPct.toFixed(2)}%</span>
        </div>
    `).join('');

    // Views Tinggi, Konversi Rendah
    const sortedV = [...list].map(d=>d.views).sort((a,b)=>a-b);
    const medianV = sortedV[Math.floor(sortedV.length/2)] || 0;
    const perhatian = [...list].filter(d=>d.views>medianV && d.konversiPct<avgKonv).sort((a,b)=>b.views-a.views).slice(0,5);
    document.getElementById('skuPerhatian').innerHTML = perhatian.length === 0
        ? '<div style="color:#bbb;font-size:0.82em;text-align:center;padding:16px;">Semua produk performa oke 👍</div>'
        : perhatian.map(d => `
            <div style="padding:7px 0;border-bottom:1px solid #f5f5f5;font-size:0.82em;">
                <div style="font-weight:600;color:#333;margin-bottom:2px;">${_skuShortName(d.sku)}</div>
                <div style="display:flex;gap:10px;color:#888;">
                    <span>👁 ${d.views.toLocaleString('id-ID')} views</span>
                    <span>📉 ${d.konversiPct.toFixed(2)}% konv</span>
                </div>
            </div>
        `).join('');

    // Repeat Buyer
    const topRepeat = [...list].filter(d=>d.repeatPct>0).sort((a,b)=>b.repeatPct-a.repeatPct).slice(0,5);
    document.getElementById('skuRepeat').innerHTML = topRepeat.length === 0
        ? '<div style="color:#bbb;font-size:0.82em;text-align:center;padding:16px;">Tersedia saat upload manual.</div>'
        : topRepeat.map((d,i) => `
            <div style="display:flex;justify-content:space-between;align-items:center;padding:6px 0;border-bottom:1px solid #f5f5f5;font-size:0.82em;">
                <div style="display:flex;align-items:center;gap:7px;">
                    <span style="background:#e0e7ff;color:#3730a3;border-radius:5px;padding:1px 7px;font-size:0.78em;font-weight:800;">#${i+1}</span>
                    <span style="color:#333;font-weight:600;">${_skuShortName(d.sku)}</span>
                </div>
                <span style="font-weight:800;color:#3730a3;white-space:nowrap;margin-left:8px;">${d.repeatPct.toFixed(1)}%</span>
            </div>
        `).join('');

    // Rekomendasi Aksi
    const rekomen = [];
    if (perhatian.length > 0) rekomen.push({ icon:'📸', color:'#f59e0b', bg:'#fffbeb', border:'#fde68a',
        judul:'Optimalkan Listing',
        isi:`<b>${perhatian.length} produk</b> views tinggi tapi konversi rendah. Cek foto, judul, dan harga.` });
    if (topRepeat.length > 0) rekomen.push({ icon:'⭐', color:'#3730a3', bg:'#eef2ff', border:'#c7d2fe',
        judul:'Andalkan Hero Product',
        isi:`<b>${_skuShortName(topRepeat[0].sku)}</b> repeat buyer tertinggi (${topRepeat[0].repeatPct.toFixed(1)}%). Prioritaskan stok & iklan.` });
    if (topOmset.length > 0) rekomen.push({ icon:'📦', color:'#166534', bg:'#f0fdf4', border:'#86efac',
        judul:'Jaga Stok Terlaris',
        isi:`<b>${_skuShortName(topOmset[0].sku)}</b> omset tertinggi (${_skuFmt(topOmset[0].omset)}). Pastikan stok cukup bulan depan.` });
    const hklo = [...list].filter(d=>d.konversiPct>avgKonv && d.omset<(totalOmset/(list.length||1))).sort((a,b)=>b.konversiPct-a.konversiPct).slice(0,1);
    if (hklo.length > 0) rekomen.push({ icon:'💹', color:'#0d9488', bg:'#f0fdfa', border:'#5eead4',
        judul:'Naikkan AOV',
        isi:`<b>${_skuShortName(hklo[0].sku)}</b> konversi bagus tapi omset kecil. Coba bundling atau naikkan harga.` });
    const hklt = [...list].filter(d=>d.keranjang>0&&d.pesanan===0).sort((a,b)=>b.keranjang-a.keranjang).slice(0,1);
    if (hklt.length > 0) rekomen.push({ icon:'🛒', color:'#ea580c', bg:'#fff7ed', border:'#fed7aa',
        judul:'Kejar Abandoned Cart',
        isi:`<b>${_skuShortName(hklt[0].sku)}</b> banyak masuk keranjang tapi belum checkout. Aktifkan voucher.` });
    if (rekomen.length < 3) rekomen.push({ icon:'📊', color:'#6b7280', bg:'#f9fafb', border:'#e5e7eb',
        judul:'Pantau Tren Bulan Ini',
        isi:'Upload data bulan berikutnya untuk melihat tren perkembangan setiap produk.' });

    document.getElementById('skuRekomen').innerHTML = rekomen.slice(0,6).map(r => `
        <div style="background:${r.bg};border:1px solid ${r.border};border-left:4px solid ${r.color};border-radius:10px;padding:13px;">
            <div style="font-size:1.1em;margin-bottom:5px;">${r.icon}</div>
            <div style="font-size:0.73em;font-weight:800;color:${r.color};text-transform:uppercase;letter-spacing:0.5px;margin-bottom:5px;">${r.judul}</div>
            <div style="font-size:0.78em;color:#555;line-height:1.6;">${r.isi}</div>
        </div>
    `).join('');

    document.getElementById('skuFileInfo').textContent = '📄 ' + fileInfo + ' — ' + list.length + ' SKU';
}


function toggleOverridePanel() {
    const panel = document.getElementById('overridePanel');
    const chevron = document.getElementById('overrideChevron');
    const isOpen = panel.style.display !== 'none';
    panel.style.display = isOpen ? 'none' : 'block';
    if (chevron) chevron.style.transform = isOpen ? '' : 'rotate(180deg)';
}

function setToggleIcon(collapsed) {
    const icon = document.getElementById('toggleIcon');
    if (!icon) return;
    // ◀ saat expanded, ▶ saat collapsed
    icon.innerHTML = collapsed
        ? '<svg width="8" height="14" viewBox="0 0 8 14" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M2 2 L6 7 L2 12" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>'
        : '<svg width="8" height="14" viewBox="0 0 8 14" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M6 2 L2 7 L6 12" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>';
}
function toggleDrawer(drawerEl) {
    drawerEl.classList.toggle('open');
}
function toggleSidebar() {
    const sidebar = document.querySelector('.sidebar');
    const isCollapsed = sidebar.classList.toggle('collapsed');
    setToggleIcon(isCollapsed);
    localStorage.setItem('sidebarCollapsed', isCollapsed ? '1' : '0');
}
(function(){
    const sidebar = document.querySelector('.sidebar');
    const isCollapsed = localStorage.getItem('sidebarCollapsed') === '1';
    if (isCollapsed && sidebar) sidebar.classList.add('collapsed');
    setToggleIcon(isCollapsed);
    if (typeof loadRekap === 'function') loadRekap();
    // Sinkronisasi panel kanan sesuai tab aktif saat load
    const activePanel = document.querySelector('.tab-panel.active');
    if (activePanel) {
        const activeName = activePanel.id.replace('tab-', '');
        const rp = document.getElementById('rightPanel');
        if (rp) rp.style.display = TABS_NO_REKOMEN.has(activeName) ? 'none' : '';
    }
})();

// ── EKSEKUSI PROGRAM ─────────────────────────────────────────
function eksekusiProgram(aksi, prog) {
    const d = window._biayaProgData || [];
    const p = d.find(x => x.prog === prog);
    if (!p) return;
    const modal = document.getElementById('eksekusiModal');
    const title = document.getElementById('eksekusiTitle');
    const body  = document.getElementById('eksekusiBody');
    if (!modal) return;

    if (aksi === 'lanjut') {
        title.innerHTML = `<span style="color:#166534;">✅ Lanjutkan ${prog}</span>`;
        body.innerHTML = `
            <div style="font-size:0.82em;color:#333;line-height:1.8;margin-bottom:14px;">
                Kamu memutuskan untuk <b>melanjutkan program ${prog}</b>.<br>
                Berikut yang perlu diperhatikan ke depan:
            </div>
            <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:14px;">
                <div style="background:#f0fdf4;border-radius:8px;padding:12px;border-left:3px solid #22c55e;">
                    <div style="font-size:0.65em;color:#166534;font-weight:800;text-transform:uppercase;margin-bottom:4px;">Target Rasio Ideal</div>
                    <div style="font-size:1.1em;font-weight:800;color:#166534;">≤ 5%</div>
                    <div style="font-size:0.65em;color:#aaa;">Rasio saat ini: ${p.rasio.toFixed(1)}%</div>
                </div>
                <div style="background:#f8f9ff;border-radius:8px;padding:12px;border-left:3px solid #3730a3;">
                    <div style="font-size:0.65em;color:#3730a3;font-weight:800;text-transform:uppercase;margin-bottom:4px;">Leverage Shopee</div>
                    <div style="font-size:1.1em;font-weight:800;color:#3730a3;">${fmtRp(p.leverRp)}</div>
                    <div style="font-size:0.65em;color:#aaa;">per Rp 1.000 seller keluar</div>
                </div>
            </div>
            <div style="background:#fffbeb;border:1px solid #fcd34d;border-radius:8px;padding:12px;font-size:0.72em;color:#92400e;line-height:1.7;">
                💡 <b>Saran:</b> Pantau rasio beban setiap bulan. Jika naik di atas 5%, evaluasi ulang apakah trafik &amp; konversi yang didapat sepadan.
                Leverage Shopee sebesar <b>${fmtRp(p.leverRp)}</b> per Rp 1.000 ${p.warn ? 'perlu dicermati — relatif kecil.' : 'sudah cukup baik — Shopee menanggung lebih banyak.'}
            </div>`;
    } else {
        title.innerHTML = `<span style="color:#dc2626;">⛔ Berhenti dari ${prog}</span>`;
        body.innerHTML = `
            <div style="font-size:0.82em;color:#333;line-height:1.8;margin-bottom:14px;">
                Kamu memutuskan untuk <b>berhenti dari program ${prog}</b>.<br>
                Pertimbangkan dampak berikut sebelum benar-benar menonaktifkan:
            </div>
            <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:14px;">
                <div style="background:#f0fdf4;border-radius:8px;padding:12px;border-left:3px solid #22c55e;">
                    <div style="font-size:0.65em;color:#166534;font-weight:800;text-transform:uppercase;margin-bottom:4px;">💰 Biaya Dihemat</div>
                    <div style="font-size:1.1em;font-weight:800;color:#166534;">+ ${fmtRp(p.ditSeller)}</div>
                    <div style="font-size:0.65em;color:#aaa;">potongan seller tidak lagi keluar</div>
                </div>
                <div style="background:#fff1f2;border-radius:8px;padding:12px;border-left:3px solid #dc2626;">
                    <div style="font-size:0.65em;color:#dc2626;font-weight:800;text-transform:uppercase;margin-bottom:4px;">📉 Omset Berisiko Hilang</div>
                    <div style="font-size:1.1em;font-weight:800;color:#dc2626;">- ${fmtRp(p.pOmset)}</div>
                    <div style="font-size:0.65em;color:#aaa;">estimasi dari ${p.nPakai} pesanan program ini</div>
                </div>
            </div>
            <div style="background:#fff9f0;border:1px solid #fed7aa;border-radius:8px;padding:12px;font-size:0.72em;color:#92400e;line-height:1.7;">
                ⚠️ <b>Perhatian:</b> Menghentikan program bisa mengurangi daya tarik toko di halaman pencarian Shopee.
                Pastikan kamu punya strategi pengganti (misalnya: naikkan harga &amp; kurangi diskon, atau aktifkan program lain) sebelum menonaktifkan.
            </div>`;
    }

    modal.style.display = 'flex';
}

function tutupEksekusiModal() {
    const modal = document.getElementById('eksekusiModal');
    if (modal) modal.style.display = 'none';
}

// ── EXPORT PDF ANALISIS BIAYA LAYANAN ────────────────────────
function exportBiayaPDF() {
    const d = window._biayaProgData || [];
    if (!d.length) { alert('Jalankan Analisis Sekarang terlebih dahulu.'); return; }

    const today = new Date().toLocaleDateString('id-ID',{day:'numeric',month:'long',year:'numeric'});
    const namaFile = (window._masterNamaToko || 'Toko Anda');

    const pageCSS = `
        @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');
        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: 'Plus Jakarta Sans', sans-serif; background: #fff; color: #333; }
        .page { width: 210mm; min-height: 297mm; padding: 18mm 16mm; page-break-after: always; }
        .page:last-child { page-break-after: avoid; }
        @media print { .no-print { display: none; } }
    `;

    const headerHTML = (prog) => `
        <div style="border-bottom:2px solid #f0f0f0;padding-bottom:14px;margin-bottom:20px;">
            <div style="font-size:0.65em;font-weight:700;color:#aaa;text-transform:uppercase;letter-spacing:1px;margin-bottom:4px;">Laporan Analisis Biaya Layanan</div>
            <div style="font-size:1.1em;font-weight:800;color:#1a1a2e;">${namaFile} — ${prog}</div>
            <div style="font-size:0.7em;color:#aaa;margin-top:2px;">Tanggal Dibuat: ${today}</div>
        </div>`;

    const glossaryPage = `
        <div class="page">
            ${headerHTML('Glosarium')}
            <div style="font-size:0.9em;font-weight:800;color:#ee4d2d;margin-bottom:16px;">Glosarium</div>
            ${[
                ['Total Pendapatan','Total pendapatan dari semua pesanan yang diproses.'],
                ['Rasio Beban Biaya','Persentase biaya program (potongan seller) terhadap total pendapatan.'],
                ['Total Voucher','Total nilai voucher yang digunakan oleh pembeli, termasuk bagian yang ditanggung Shopee dan Seller.'],
                ['Potongan Seller','Total biaya layanan yang dibebankan kepada kamu untuk program ini.'],
                ['Ditanggung Shopee','Total nilai subsidi atau bagian voucher yang ditanggung oleh Shopee.'],
                ['Leverage Shopee','Seberapa besar subsidi Shopee untuk setiap Rp 1.000 yang kamu keluarkan.'],
                ['Rata-Rata Voucher','Rata-rata nilai voucher yang didapat oleh pembeli yang menggunakan program ini.'],
            ].map(([t,d]) => `
                <div style="margin-bottom:14px;padding-bottom:14px;border-bottom:1px solid #f5f5f5;">
                    <div style="font-size:0.85em;font-weight:800;color:#1a1a2e;margin-bottom:3px;">${t}</div>
                    <div style="font-size:0.78em;color:#666;line-height:1.7;">${d}</div>
                </div>`).join('')}
            <div style="margin-top:40px;padding-top:16px;border-top:2px solid #f0f0f0;font-size:0.68em;color:#aaa;text-align:center;line-height:1.8;">
                Burhan Alfironi Muktamar (Burhanmology) — Seller Mentor Shopee<br>
                IG: @burhanmology | Tiktok: @burhanmology | Youtube: Burhanmology | WA Admin: +62 851-4296-4667
            </div>
        </div>`;

    const progPages = d.map(p => {
        const accentCol = {'Vchr Extra':'#3730a3','Video Extra':'#0d9488','Live Extra':'#dc2626','GOFS Extra':'#0369a1','Voucher Penjual':'#92400e','Tanpa Program':'#6b7280'}[p.prog] || '#3730a3';
        const rasioLabel = p.warn ? '⚠️ Di atas batas wajar (>5%)' : '✅ Masih dalam batas wajar';
        const rasioColor = p.warn ? '#dc2626' : '#166534';
        return `
        <div class="page">
            ${headerHTML('Analisis ' + p.prog)}
            <!-- HERO STATS -->
            <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px;margin-bottom:20px;">
                <div style="border:1.5px solid #e8e8e8;border-radius:10px;padding:14px;text-align:center;border-top:4px solid ${accentCol};">
                    <div style="font-size:0.6em;color:#888;font-weight:700;text-transform:uppercase;margin-bottom:4px;">Total Pendapatan</div>
                    <div style="font-size:1.1em;font-weight:800;color:#1a1a2e;">${fmtRp(p.totalOmset)}</div>
                    <div style="font-size:0.62em;color:#aaa;">${p.n} pesanan</div>
                </div>
                <div style="border:1.5px solid #e8e8e8;border-radius:10px;padding:14px;text-align:center;border-top:4px solid ${p.warn?'#dc2626':'#22c55e'};">
                    <div style="font-size:0.6em;color:#888;font-weight:700;text-transform:uppercase;margin-bottom:4px;">Rasio Beban Biaya</div>
                    <div style="font-size:1.4em;font-weight:800;color:${rasioColor};">${p.rasio.toFixed(1)}%</div>
                    <div style="font-size:0.62em;color:#aaa;">Biaya / Pendapatan</div>
                </div>
                <div style="border:1.5px solid #e8e8e8;border-radius:10px;padding:14px;text-align:center;border-top:4px solid ${accentCol};">
                    <div style="font-size:0.6em;color:#888;font-weight:700;text-transform:uppercase;margin-bottom:4px;">Total Voucher Program</div>
                    <div style="font-size:1.1em;font-weight:800;color:#1a1a2e;">${fmtRp(p.totalVchr)}</div>
                    <div style="font-size:0.62em;color:#aaa;">${p.totalOmset>0?(p.totalVchr/p.totalOmset*100).toFixed(1):'0'}% dari total</div>
                </div>
            </div>
            <!-- DETAIL STATS -->
            <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px;margin-bottom:20px;">
                <div style="background:#f8f9ff;border-radius:8px;padding:12px;text-align:center;">
                    <div style="font-size:0.6em;color:#888;font-weight:700;text-transform:uppercase;margin-bottom:3px;">Rata-Rata Voucher</div>
                    <div style="font-size:0.95em;font-weight:800;color:#1a1a2e;">${fmtRp(p.rataVchr)}</div>
                    <div style="font-size:0.6em;color:#aaa;">${p.nPakai} pesanan (${p.pctPakai}%)</div>
                </div>
                <div style="background:#fff1f2;border-radius:8px;padding:12px;text-align:center;">
                    <div style="font-size:0.6em;color:#be123c;font-weight:700;text-transform:uppercase;margin-bottom:3px;">Potongan Seller</div>
                    <div style="font-size:0.95em;font-weight:800;color:#9f1239;">${fmtRp(p.ditSeller)}</div>
                    <div style="font-size:0.6em;color:#aaa;">${p.pctSl.toFixed(1)}% dari total voucher</div>
                </div>
                <div style="background:#f0f9ff;border-radius:8px;padding:12px;text-align:center;">
                    <div style="font-size:0.6em;color:#0369a1;font-weight:700;text-transform:uppercase;margin-bottom:3px;">Ditanggung Shopee</div>
                    <div style="font-size:0.95em;font-weight:800;color:#075985;">${fmtRp(p.ditShopee)}</div>
                    <div style="font-size:0.6em;color:#aaa;">${p.pctSh.toFixed(1)}% dari total voucher</div>
                </div>
            </div>
            <!-- RINGKASAN -->
            <div style="background:#f8faff;border-radius:10px;padding:16px;border-left:4px solid ${accentCol};margin-bottom:16px;">
                <div style="font-size:0.7em;font-weight:800;color:${accentCol};text-transform:uppercase;letter-spacing:0.5px;margin-bottom:8px;">Ringkasan Analisis</div>
                <div style="font-size:0.78em;line-height:1.85;color:#555;">
                    Dari total <b>${p.n} pesanan</b>, sebanyak <b>${p.nPakai} pesanan (${p.pctPakai}%)</b> menggunakan voucher program <b>${p.prog}</b>.
                    Rata-rata, pembeli mendapatkan potongan sebesar <b>${fmtRp(p.rataVchr)}</b> per pesanan.<br><br>
                    Rasio beban biaya program ini terhadap total pendapatan adalah sebesar <b>${p.rasio.toFixed(2)}%</b>.
                    Dari total voucher yang terpakai, <b>${p.pctSh.toFixed(1)}%</b> disubsidi oleh Shopee, sementara <b>${p.pctSl.toFixed(1)}%</b> menjadi biaya yang ditanggung seller.
                    Artinya, untuk setiap <b>Rp 1.000</b> yang kamu keluarkan, Shopee memberikan dukungan subsidi sebesar <b>${fmtRp(p.leverRp)}</b>.
                </div>
                <div style="margin-top:10px;font-size:0.72em;font-weight:700;color:${rasioColor};">${rasioLabel}</div>
            </div>
            <!-- SARAN -->
            <div style="background:#fffbeb;border-radius:10px;padding:14px;border-left:4px solid #f59e0b;margin-bottom:20px;">
                <div style="font-size:0.7em;font-weight:800;color:#92400e;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:6px;">Saran Pertimbangan</div>
                <div style="font-size:0.75em;line-height:1.85;color:#78350f;">
                    Analisis ini bertujuan untuk memberikan gambaran. Keputusan untuk lanjut atau berhenti dari program ada di tangan kamu.
                    Pertimbangkan apakah benefit dari program (potensi kenaikan trafik &amp; konversi) sepadan dengan rasio biaya yang dikeluarkan.
                    Jika memutuskan berhenti, pertimbangkan juga potensi penurunan jumlah pesanan karena hilangnya daya tarik promo di toko kamu.
                </div>
            </div>
            <div style="padding-top:14px;border-top:1.5px solid #f0f0f0;font-size:0.65em;color:#aaa;text-align:center;line-height:1.8;">
                Burhan Alfironi Muktamar (Burhanmology) — Seller Mentor Shopee<br>
                IG: @burhanmology | Tiktok: @burhanmology | Youtube: Burhanmology | WA Admin: +62 851-4296-4667
            </div>
        </div>`;
    }).join('');

    const fullHTML = `<!DOCTYPE html><html lang="id"><head>
        <meta charset="UTF-8"><title>Laporan Analisis Biaya Layanan — ${namaFile}</title>
        <style>${pageCSS}</style></head><body>
        ${progPages}${glossaryPage}
        <div class="no-print" style="position:fixed;bottom:20px;right:20px;">
            <button onclick="window.print()" style="background:#ee4d2d;color:#fff;border:none;padding:10px 20px;border-radius:8px;font-size:0.85em;font-weight:800;cursor:pointer;">🖨️ Print / Save PDF</button>
        </div>
    </body></html>`;

    const blob = new Blob([fullHTML], {type:'text/html'});
    const url = URL.createObjectURL(blob);
    const win = window.open(url, '_blank');
    if (win) win.focus();
}

