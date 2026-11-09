/* =========================================================
   SIMPATI — Manajemen Aset Tetap & Perpindahan Barang
   Aplikasi statis: data aset dari Excel, mutasi di localStorage
   ========================================================= */
(function () {
  'use strict';

  /* ---------------- Data & konstanta ---------------- */
  var DEFAULT_DB = window.ASSET_DB;
  if (!DEFAULT_DB) { document.body.innerHTML = '<p style="padding:24px">Data aset tidak ditemukan.</p>'; return; }

  var DB = null;
  var LOC = [];
  var SUB = [];
  var KEL = [];
  var BRAND = [];
  var ASET = [];
  var KATEGORI = window.ASSET_KATEGORI || { list: [], life: [], map: {} };
  var KAT_LIST = Array.isArray(KATEGORI.list) ? KATEGORI.list : [];
  var KAT_LIFE = Array.isArray(KATEGORI.life) ? KATEGORI.life : [];
  var KAT_MAP = KATEGORI.map || {};
  var LOC_BY_CODE = {};
  var DATASET = {
    kind: 'default',
    source: (DEFAULT_DB.meta && DEFAULT_DB.meta.source) || 'Data Aset',
    fileName: '',
    sheet: 'ALL',
    updatedAt: '',
    importedAt: ''
  };

  function buildData(db) {
    DB = db || { aset: [], lokasi: [], sub: [], kel: [], merek: [], meta: {} };
    LOC = Array.isArray(DB.lokasi) ? DB.lokasi : [];
    SUB = Array.isArray(DB.sub) ? DB.sub : [];
    KEL = Array.isArray(DB.kel) ? DB.kel : [];
    BRAND = Array.isArray(DB.merek) ? DB.merek : [];
    ASET = (Array.isArray(DB.aset) ? DB.aset : []).filter(Array.isArray).map(function (a) {
      return {
        ref: String(a[0] == null ? '' : a[0]),
        name: String(a[1] == null ? '' : a[1]),
        loc: a[2] == null ? -1 : a[2],
        date: a[3] || '',
        sub: a[4] == null ? -1 : a[4],
        brand: a[5] == null ? -1 : a[5],
        price: +a[6] || 0,
        kel: a[7] == null ? -1 : a[7],
        desc: a[8] || ''
      };
    });
    LOC_BY_CODE = {};
    LOC.forEach(function (l, i) { if (l && l[0]) LOC_BY_CODE[l[0]] = i; });
  }

  var SUB_COLOR = {
    'ALAT MEDIS DAN KESEHATAN': '#0f9d8b',
    'ALAT UMUM DAN RUMAH TANGGA': '#3b82f6',
    'PERALATAN KOMPUTER': '#8b5cf6',
    'KENDARAAN BERMOTOR': '#e8a33d',
    'PERANGKAT LUNAK': '#e06a9a',
    'LAIN-LAIN': '#94a9b8'
  };
  var SUB_BADGE = {
    'ALAT MEDIS DAN KESEHATAN': 'badge-teal',
    'ALAT UMUM DAN RUMAH TANGGA': 'badge-blue',
    'PERALATAN KOMPUTER': 'badge-violet',
    'KENDARAAN BERMOTOR': 'badge-amber',
    'PERANGKAT LUNAK': 'badge-pink',
    'LAIN-LAIN': 'badge-gray'
  };

  var STORAGE_KEY = 'simpati_state_v1';
  var REPAIR_KEY = 'simpati_repairs_v1';
  var ADMIN_PIN_KEY = 'simpati_admin_pin_v1';
  var ADMIN_LOG_KEY = 'simpati_admin_log_v1';
  var ADMIN_SESSION_KEY = 'simpati_admin_session';
  var REPAIR_THRESHOLD = 4000000;
  var REPAIR_PERCENT = 0.15;
  var KAT_BADGE = ['kat-alum1', 'kat-alum2', 'kat-alkes1', 'kat-alkes2', 'kat-mobil', 'kat-motor', 'kat-kso1', 'kat-kso2'];
  var KAT_DOT = ['#2f7de1', '#5b8def', '#0f9d8b', '#34b3a0', '#e8a33d', '#d97706', '#8b5cf6', '#a78bfa'];

  var REASONS = ['Relokasi ruang', 'Kebutuhan pelayanan', 'Perbaikan / servis', 'Peminjaman unit lain', 'Penggantian unit', 'Lainnya'];
  var ACR = {
    IGD: 1, ICU: 1, NICU: 1, PICU: 1, VK: 1, OK: 1, RSH: 1, TPS: 1, CSSU: 1, EMR: 1,
    MRI: 1, CT: 1, USG: 1, KBBL: 1, RO: 1, EDP: 1, PKRS: 1, KTK: 1, VIP: 1, THT: 1,
    OBSGYN: 1, IP: 1, CCTV: 1, UPS: 1, FCC: 1, KSO: 1, BA: 1, GL: 1, IGD: 1, HD: 1
  };

  /* ---------------- Util ---------------- */
  var nf = new Intl.NumberFormat('id-ID');
  var dfShort = new Intl.DateTimeFormat('id-ID', { day: '2-digit', month: 'short', year: 'numeric' });
  var dfLong = new Intl.DateTimeFormat('id-ID', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });

  function $(sel, root) { return (root || document).querySelector(sel); }
  function $$(sel, root) { return Array.prototype.slice.call((root || document).querySelectorAll(sel)); }
  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }
  function rp(n) { return 'Rp ' + nf.format(n || 0); }
  function moneyShort(n) {
    if (n >= 1e12) return { v: (n / 1e12).toFixed(2).replace('.', ','), u: 'triliun' };
    if (n >= 1e9) return { v: (n / 1e9).toFixed(2).replace('.', ','), u: 'miliar' };
    if (n >= 1e6) return { v: (n / 1e6).toFixed(1).replace('.', ','), u: 'juta' };
    return { v: nf.format(n), u: '' };
  }
  function dateObj(iso) {
    if (!iso) return null;
    var p = String(iso).slice(0, 10).split('-');
    return new Date(+p[0], +p[1] - 1, +p[2]);
  }
  function fmtDate(iso) {
    var d = dateObj(iso);
    return d ? dfShort.format(d) : '—';
  }
  function todayISO() {
    var d = new Date();
    return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
  }
  function thisMonth(iso) { return String(iso || '').slice(0, 7) === todayISO().slice(0, 7); }
  function titleCase(s) {
    return String(s || '').split(/\s+/).map(function (w) {
      if (!w) return w;
      if (ACR[w.toUpperCase()] || /^[0-9.\-/,]+$/.test(w)) return w;
      return w.charAt(0).toUpperCase() + w.slice(1).toLowerCase();
    }).join(' ');
  }
  function locRaw(id) { return LOC[id] || ['', 'TANPA LOKASI']; }
  function locCode(id) { return locRaw(id)[0]; }
  function locName(id) {
    return titleCase(locRaw(id)[1].replace(/^RSH MADIUN\s+/i, '').trim());
  }
  function locFull(id) {
    var c = locCode(id), n = locName(id);
    return (c ? c + ' · ' : '') + n;
  }
  function curLoc(a) {
    var ov = state.overrides[a.ref];
    return (ov === undefined || ov === null) ? a.loc : ov;
  }
  function brandName(a) { return a.brand >= 0 && BRAND[a.brand] ? BRAND[a.brand] : ''; }
  function subName(id) { return SUB[id] || 'LAIN-LAIN'; }
  function katIndexOf(a) {
    var v = KAT_MAP[a.ref];
    return (typeof v === 'number' && KAT_LIST[v]) ? v : -1;
  }
  function katNameOf(a) {
    var i = katIndexOf(a);
    return i >= 0 ? KAT_LIST[i] : 'Belum berkategori';
  }
  function katLifeOf(a) {
    var i = katIndexOf(a);
    return i >= 0 ? (+KAT_LIFE[i] || 0) : 0;
  }
  function rupiahShortText(n) { var m = moneyShort(n); return 'Rp ' + m.v + (m.u ? ' ' + m.u : ''); }

  function fmtDateTime(iso) {
    if (!iso) return '—';
    var d = new Date(iso);
    if (isNaN(d.getTime())) return '—';
    return d.toLocaleDateString('id-ID', { day: '2-digit', month: 'short', year: 'numeric' }) + ' · ' +
      d.toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' });
  }

  function uniqueLocations() {
    var map = {};
    ASET.forEach(function (a) { map[curLoc(a)] = 1; });
    return Object.keys(map).length;
  }

  function datasetSummary(db) {
    var total = 0;
    var locs = {};
    var subs = {};
    (db.aset || []).forEach(function (a) {
      total += +a[6] || 0;
      if (a[2] != null) locs[a[2]] = 1;
      if (a[4] != null) subs[a[4]] = 1;
    });
    return { total: total, count: (db.aset || []).length, locations: Object.keys(locs).length, subs: Object.keys(subs).length };
  }

  function withMeta(db, info) {
    var s = datasetSummary(db);
    var when = (info && info.updatedAt) || new Date().toISOString();
    db.meta = {
      total: s.total,
      count: s.count,
      locations: s.locations,
      source: (info && info.source) || (db.meta && db.meta.source) || 'Data Aset',
      created: String(when).slice(0, 10),
      sheet: (info && info.sheet) || '',
      kind: (info && info.kind) || 'custom',
      fileName: (info && info.fileName) || '',
      updatedAt: when
    };
    return db;
  }

  function cloneDb(db) {
    return {
      meta: db.meta ? JSON.parse(JSON.stringify(db.meta)) : {},
      lokasi: (db.lokasi || []).slice(),
      sub: (db.sub || []).slice(),
      kel: (db.kel || []).slice(),
      merek: (db.merek || []).slice(),
      aset: (db.aset || []).slice()
    };
  }

  function cleanText(v) {
    return String(v == null ? '' : v).replace(/\s+/g, ' ').trim();
  }

  function parseMoneyValue(v) {
    if (typeof v === 'number') return isFinite(v) ? Math.round(v) : 0;
    var s = String(v == null ? '' : v).replace(/[^\d,.-]/g, '');
    if (!s) return 0;
    var lastComma = s.lastIndexOf(',');
    var lastDot = s.lastIndexOf('.');
    if (lastComma >= 0 && lastDot >= 0) {
      if (lastComma > lastDot) s = s.replace(/\./g, '').replace(',', '.');
      else s = s.replace(/,/g, '');
    } else if (lastComma >= 0) {
      if (s.length - lastComma - 1 <= 2) s = s.replace(/,/g, '.');
      else s = s.replace(/,/g, '');
    }
    var n = parseFloat(s);
    return isFinite(n) ? Math.round(n) : 0;
  }

  function serialToISO(serial) {
    if (!isFinite(serial) || serial < 1 || serial > 2958465) return '';
    var ms = Math.round((serial - 25569) * 86400000);
    var d = new Date(ms);
    if (isNaN(d.getTime())) return '';
    return d.getUTCFullYear() + '-' + String(d.getUTCMonth() + 1).padStart(2, '0') + '-' + String(d.getUTCDate()).padStart(2, '0');
  }

  function parseDateValue(v) {
    if (v == null || v === '') return '';
    if (typeof v === 'number') return serialToISO(v);
    var s = String(v).trim();
    if (!s) return '';
    var iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (iso) return iso[1] + '-' + iso[2] + '-' + iso[3];
    var ymd = s.match(/^(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})$/);
    if (ymd) {
      return ymd[1] + '-' + String(+ymd[2]).padStart(2, '0') + '-' + String(+ymd[3]).padStart(2, '0');
    }
    var dmy = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/);
    if (dmy) {
      var a = +dmy[1], b = +dmy[2];
      if (a > 12 && b <= 12) { var tmp = a; a = b; b = tmp; }
      if (a >= 1 && a <= 31 && b >= 1 && b <= 12) {
        return dmy[3] + '-' + String(b).padStart(2, '0') + '-' + String(a).padStart(2, '0');
      }
    }
    if (/^\d{5}(\.\d+)?$/.test(s)) {
      var serial = parseFloat(s);
      if (serial > 20000) return serialToISO(serial);
    }
    var d = new Date(s);
    if (!isNaN(d.getTime())) {
      return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
    }
    return '';
  }

  function locDisplayName(id) {
    if (id == null || id < 0 || !LOC[id]) return '—';
    return locName(id);
  }

  function toast(msg) {
    var el = $('#toast');
    el.innerHTML = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M5 12.5 10 17l9-10"/></svg><span>' + esc(msg) + '</span>';
    el.classList.add('is-show');
    clearTimeout(toast._t);
    toast._t = setTimeout(function () { el.classList.remove('is-show'); }, 3200);
  }

  /* ---------------- State ---------------- */
  var state = { transfers: [], overrides: {} };
  var repairs = [];
  var admin = { unlocked: false, log: [], pendingRestore: null };
  var upload = {
    step: 1, fileName: '', fileSize: 0, fileType: '', sheets: [], sheetIndex: 0,
    sheetName: '', rows: [], headerRow: 1, headers: [], map: {}, built: null,
    mode: 'replace', pending: null, success: null, error: ''
  };
  var ui = {
    view: 'ringkasan',
    page: 1, pageSize: 12, search: '', sub: '', loc: '', sort: 'nama',
    histSearch: '', histReason: '',
    current: null, pickerIndex: -1,
    repairAsset: null, repairPickerIndex: -1,
    depSearch: '', depKat: '', depStatus: '', depSort: 'nama', depPage: 1, depPageSize: 12,
    pendingDelete: null
  };

  function save() {
    try { localStorage.setItem(STORAGE_KEY, JSON.stringify(state)); } catch (e) { /* mode privat */ }
  }
  function load() {
    try {
      var raw = localStorage.getItem(STORAGE_KEY);
      if (raw) {
        var p = JSON.parse(raw);
        if (p && Array.isArray(p.transfers)) {
          state.transfers = p.transfers;
          state.overrides = p.overrides || {};
          return true;
        }
      }
    } catch (e) { /* abaikan */ }
    return false;
  }

  /* ---------------- Seed contoh (dapat di-reset) ---------------- */
  function findAsset(kw) {
    for (var i = 0; i < ASET.length; i++) {
      if (ASET[i].name.toUpperCase().indexOf(kw) === 0) return ASET[i];
    }
    for (var j = 0; j < ASET.length; j++) {
      if (ASET[j].name.toUpperCase().indexOf(kw) !== -1) return ASET[j];
    }
    return ASET[0];
  }

  function seedExamples() {
    var specs = [
      { kw: 'KURSI TUNGGU', code: '049.01.02.20', date: '2026-10-27', pic: 'Rina Kusuma', reason: 'Relokasi ruang', note: 'Berita acara BA-102/X/2026' },
      { kw: 'LEMARI', code: '049.05.26.18', date: '2026-10-29', pic: 'Bagus Prakoso', reason: 'Kebutuhan pelayanan', note: 'Penambahan ruang arsip rekam medis' },
      { kw: 'TENSIMETER', code: '049.03.20.14', date: '2026-11-01', pic: 'Dewi Anggraini', reason: 'Peminjaman unit lain', note: 'Dipinjam untuk poli kamala' },
      { kw: 'KOMPUTER', code: '049.05.28.06', date: '2026-11-02', pic: 'Yudi Hartono', reason: 'Penggantian unit', note: 'Unit lama masuk gudang IT' },
      { kw: 'TROLLEY', code: '049.01.04.02', date: '2026-11-03', pic: 'Sari Wulandari', reason: 'Perbaikan / servis', note: 'Kalibrasi di IGD' }
    ];
    var list = [];
    specs.forEach(function (s, i) {
      var a = findAsset(s.kw);
      var to = LOC_BY_CODE[s.code];
      if (to === undefined) to = (a.loc + 1) % LOC.length;
      if (to === a.loc) to = (a.loc + 2) % LOC.length;
      list.push({
        id: 'SEED-' + (i + 1),
        ref: a.ref,
        name: a.name,
        from: a.loc,
        to: to,
        date: s.date,
        pic: s.pic,
        reason: s.reason,
        note: s.note,
        createdAt: s.date + 'T09:0' + i + ':00.000Z',
        demo: true
      });
    });
    state.transfers = list;
    rebuildOverrides();
  }

  function rebuildOverrides() {
    var o = {};
    state.transfers.slice().sort(function (a, b) {
      return String(a.date + a.createdAt).localeCompare(String(b.date + b.createdAt));
    }).forEach(function (t) { o[t.ref] = t.to; });
    state.overrides = o;
  }

  /* ---------------- Navigasi ---------------- */
  var TITLES = { ringkasan: 'Ringkasan', aset: 'Data Aset', perpindahan: 'Perpindahan', perbaikan: 'Perbaikan', penyusutan: 'Penyusutan', riwayat: 'Riwayat', tambah: 'Tambah Barang', admin: 'Admin' };

  function go(view) {
    if (!TITLES[view]) return;
    ui.view = view;
    $$('.view').forEach(function (s) { s.classList.toggle('is-active', s.id === 'view-' + view); });
    $$('.nav-item').forEach(function (b) { b.classList.toggle('is-active', b.getAttribute('data-view') === view); });
    $('#page-title').textContent = TITLES[view];
    document.title = 'SIMPATI — ' + TITLES[view];
    closeSidebar();
    window.scrollTo({ top: 0, behavior: 'smooth' });

    if (view === 'ringkasan') renderDashboard();
    if (view === 'aset') renderAset();
    if (view === 'perpindahan') updatePreview();
    if (view === 'perbaikan') { renderRepair(); renderRepairs(); }
    if (view === 'penyusutan') renderPenyusutan();
    if (view === 'tambah') renderAddPreview();
    if (view === 'admin') renderAdmin();
  }

  function openSidebar() { $('#sidebar').classList.add('is-open'); $('#overlay').hidden = false; }
  function closeSidebar() { $('#sidebar').classList.remove('is-open'); $('#overlay').hidden = true; }

  /* ---------------- Dashboard ---------------- */
  function groupBySub() {
    var map = {};
    SUB.forEach(function (s, i) { map[i] = { name: s, count: 0, value: 0 }; });
    ASET.forEach(function (a) {
      if (!map[a.sub]) map[a.sub] = { name: subName(a.sub), count: 0, value: 0 };
      map[a.sub].count++;
      map[a.sub].value += a.price;
    });
    return Object.keys(map).map(function (k) {
      return { id: +k, name: map[k].name, count: map[k].count, value: map[k].value };
    }).filter(function (x) { return x.count > 0; });
  }

  function moveCountMonth() {
    var refs = {};
    var n = 0;
    state.transfers.forEach(function (t) { if (thisMonth(t.date)) { n++; refs[t.ref] = 1; } });
    return { n: n, refs: Object.keys(refs).length };
  }

  function renderDashboard() {
    var totalValue = DB.meta.total || ASET.reduce(function (s, a) { return s + a.price; }, 0);
    var locCount = {};
    ASET.forEach(function (a) { locCount[curLoc(a)] = (locCount[curLoc(a)] || 0) + 1; });
    var nLoc = Object.keys(locCount).length;
    var mm = moveCountMonth();
    var avg = totalValue / (ASET.length || 1);
    var m = moneyShort(totalValue);
    var ma = moneyShort(avg);

    $('#kpi-grid').innerHTML = [
      kpi('Total Nilai Aset', 'Rp ' + m.v, m.u, 'wallet', '#0d9488', '#e3f6f2',
        nf.format(ASET.length) + ' unit · ' + groupBySub().length + ' sub kelompok'),
      kpi('Jumlah Unit Aset', nf.format(ASET.length), 'unit', 'layers', '#2f7de1', '#e9f1fd',
        'Tersebar di ' + nf.format(nLoc) + ' lokasi / ruang'),
      kpi('Perpindahan Bulan Ini', nf.format(mm.n), 'kali', 'swap', '#8b5cf6', '#eeecfe',
        mm.refs + ' aset berbeda dipindahkan'),
      kpi('Rata-rata Nilai Aset', 'Rp ' + ma.v, ma.u, 'trend', '#e59a17', '#fdf3de',
        'Per unit berdasarkan harga perolehan')
    ].join('');

    renderSubChart();
    renderDonut();
    renderRecent();
    renderTopLocations();

    $('#nav-count-aset').textContent = nf.format(ASET.length);
    $('#nav-count-riwayat').textContent = state.transfers.length || '';
    $('#source-meta').textContent = nf.format(ASET.length) + ' unit · ' + nf.format(nLoc) + ' lokasi';
    $('#chart-total-badge').textContent = rupiahShortText(totalValue);
  }

  function kpi(label, value, unit, icon, color, tint, foot) {
    var paths = {
      wallet: '<path d="M4 7.5A2.5 2.5 0 0 1 6.5 5H19a1 1 0 0 1 1 1v12a1 1 0 0 1-1 1H6.5A2.5 2.5 0 0 1 4 16.5v-9Z"/><path d="M4 8h13"/><path d="M16 12.5h2"/>',
      layers: '<path d="m12 3 8 4.5-8 4.5-8-4.5L12 3Z"/><path d="m4 12 8 4.5 8-4.5"/><path d="m4 16.5 8 4.5 8-4.5"/>',
      swap: '<path d="M4 7h13l-3-3M20 17H7l3 3"/><path d="M17 7v4M7 17v-4"/>',
      trend: '<path d="M4 17l5-5 3.5 3L20 7"/><path d="M15 7h5v5"/>',
      calendar: '<rect x="3" y="5" width="18" height="16" rx="2"/><path d="M8 3v4M16 3v4M3 10h18"/>'
    };
    return '<article class="kpi" style="--kpi-color:' + color + ';--kpi-tint:' + tint + '">' +
      '<div class="kpi-top"><span class="kpi-label">' + esc(label) + '</span>' +
      '<span class="kpi-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round">' + paths[icon] + '</svg></span></div>' +
      '<div class="kpi-value">' + esc(value) + (unit ? ' <small>' + esc(unit) + '</small>' : '') + '</div>' +
      '<div class="kpi-foot"><span class="dot"></span>' + esc(foot) + '</div>' +
      '</article>';
  }

  function renderSubChart() {
    var groups = groupBySub().sort(function (a, b) { return b.value - a.value; });
    var max = Math.max.apply(null, groups.map(function (g) { return g.value; })) || 1;
    $('#chart-sub').innerHTML = groups.map(function (g) {
      var c = SUB_COLOR[g.name] || '#94a9b8';
      var pct = Math.max(2, (g.value / max) * 100);
      return '<div class="bar-row">' +
        '<div class="bar-label"><span class="swatch" style="background:' + c + '"></span>' + esc(titleCase(g.name)) + '</div>' +
        '<div class="bar-value">' + rupiahShortText(g.value) + ' <span style="color:#5b7285;font-weight:500">(' + nf.format(g.count) + ')</span></div>' +
        '<div class="bar-track"><div class="bar-fill" style="width:' + pct + '%;--c:' + c + '"></div></div>' +
        '</div>';
    }).join('');
  }

  function renderDonut() {
    var groups = groupBySub().sort(function (a, b) { return b.count - a.count; });
    var total = groups.reduce(function (s, g) { return s + g.count; }, 0) || 1;
    var acc = 0;
    var stops = groups.map(function (g) {
      var start = (acc / total) * 100;
      acc += g.count;
      var end = (acc / total) * 100;
      var c = SUB_COLOR[g.name] || '#94a9b8';
      return c + ' ' + start.toFixed(2) + '% ' + end.toFixed(2) + '%';
    });
    $('#chart-donut').style.background = 'conic-gradient(' + stops.join(',') + ')';
    $('#donut-total').textContent = nf.format(total);
    $('#donut-legend').innerHTML = groups.map(function (g) {
      var c = SUB_COLOR[g.name] || '#94a9b8';
      return '<div class="legend-item"><span class="swatch" style="background:' + c + '"></span>' +
        '<span class="legend-name">' + esc(titleCase(g.name)) + '</span>' +
        '<span class="legend-val">' + nf.format(g.count) + '</span></div>';
    }).join('');
  }

  function sortedTransfers() {
    return state.transfers.slice().sort(function (a, b) {
      return String(b.date + b.createdAt).localeCompare(String(a.date + a.createdAt));
    });
  }

  function renderRecent() {
    var list = sortedTransfers().slice(0, 5);
    if (!list.length) {
      $('#recent-moves').innerHTML = '<div class="preview-empty"><div class="preview-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.4"><path d="M4 7h13l-3-3M20 17H7l3 3"/><path d="M17 7v4M7 17v-4"/></svg></div><p>Belum ada perpindahan tercatat.</p></div>';
      return;
    }
    $('#recent-moves').innerHTML = list.map(function (t) {
      return '<button class="move-row" data-ref="' + esc(t.ref) + '" style="background:none;border:0;width:100%;text-align:left">' +
        '<span class="move-thumb"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M4 7h13l-3-3M20 17H7l3 3"/><path d="M17 7v4M7 17v-4"/></svg></span>' +
        '<span class="move-main"><span class="move-title">' + esc(t.name) + '</span>' +
        '<span class="move-meta">' + esc(locName(t.from)) + '<span class="arrow">→</span>' + esc(locName(t.to)) + '</span></span>' +
        '<span class="move-date">' + fmtDate(t.date) + '</span></button>';
    }).join('');
  }

  function renderTopLocations() {
    var map = {};
    ASET.forEach(function (a) { var l = curLoc(a); map[l] = (map[l] || 0) + 1; });
    var arr = Object.keys(map).map(function (k) { return { loc: +k, n: map[k] }; })
      .sort(function (a, b) { return b.n - a.n; }).slice(0, 5);
    var max = arr.length ? arr[0].n : 1;
    $('#top-locs').innerHTML = arr.map(function (r, i) {
      return '<div class="rank-item' + (i === 0 ? ' top' : '') + '">' +
        '<span class="rank-no">' + (i + 1) + '</span>' +
        '<span class="rank-name" title="' + esc(locFull(r.loc)) + '">' + esc(locName(r.loc)) + '</span>' +
        '<span class="rank-count">' + nf.format(r.n) + '</span>' +
        '<span class="rank-track"><span class="rank-fill" style="display:block;width:' + Math.max(6, (r.n / max) * 100) + '%"></span></span>' +
        '</div>';
    }).join('');
  }

  /* ---------------- Data aset ---------------- */
  function buildFilters() {
    var subSel = $('#aset-sub');
    subSel.innerHTML = '<option value="">Semua sub kelompok</option>' + SUB.map(function (s, i) {
      return '<option value="' + i + '">' + esc(titleCase(s)) + '</option>';
    }).join('');

    var locOrder = LOC.map(function (l, i) { return { i: i, code: l[0], name: l[1] }; })
      .sort(function (a, b) { return String(a.code).localeCompare(String(b.code)); });
    $('#aset-lokasi').innerHTML = '<option value="">Semua lokasi</option>' + locOrder.map(function (l) {
      return '<option value="' + l.i + '">' + esc((l.code ? l.code + ' · ' : '') + titleCase(l.name.replace(/^RSH MADIUN\s+/i, ''))) + '</option>';
    }).join('');

    var toSel = $('#move-to');
    toSel.innerHTML = '<option value="">Pilih lokasi tujuan…</option>' + locOrder.map(function (l) {
      return '<option value="' + l.i + '">' + esc((l.code ? l.code + ' · ' : '') + titleCase(l.name.replace(/^RSH MADIUN\s+/i, ''))) + '</option>';
    }).join('');

    var addSub = $('#add-sub');
    if (addSub) {
      addSub.innerHTML = '<option value="">Pilih sub kelompok…</option>' + SUB.map(function (s, i) {
        return '<option value="' + i + '">' + esc(titleCase(s)) + '</option>';
      }).join('');
    }
    var addLoc = $('#add-loc');
    if (addLoc) {
      addLoc.innerHTML = '<option value="">Pilih lokasi…</option>' + locOrder.map(function (l) {
        return '<option value="' + l.i + '">' + esc((l.code ? l.code + ' · ' : '') + titleCase(l.name.replace(/^RSH MADIUN\s+/i, ''))) + '</option>';
      }).join('');
    }

    var depKat = $('#dep-kat');
    if (depKat) {
      var counts = depKatCounts();
      depKat.innerHTML = '<option value="">Semua kategori</option>' + KAT_LIST.map(function (name, i) {
        return '<option value="' + i + '">' + esc(name) + ' (' + nf.format(counts.list[i] || 0) + ')</option>';
      }).join('') + (counts.unknown ? '<option value="unknown">Belum berkategori (' + nf.format(counts.unknown) + ')</option>' : '');
    }
  }

  function filteredAssets() {
    var q = ui.search.trim().toLowerCase();
    var sub = ui.sub === '' ? -1 : +ui.sub;
    var loc = ui.loc === '' ? -1 : +ui.loc;

    var rows = ASET.filter(function (a) {
      if (sub >= 0 && a.sub !== sub) return false;
      if (loc >= 0 && curLoc(a) !== loc) return false;
      if (!q) return true;
      var hay = a.name + ' ' + a.ref + ' ' + brandName(a) + ' ' + a.desc + ' ' + locRaw(a.loc)[1] + ' ' + locRaw(curLoc(a))[1] + ' ' + locCode(curLoc(a));
      return hay.toLowerCase().indexOf(q) !== -1;
    });

    var s = ui.sort;
    rows.sort(function (a, b) {
      if (s === 'nilai-desc') return b.price - a.price;
      if (s === 'nilai-asc') return a.price - b.price;
      if (s === 'baru') return String(b.date).localeCompare(String(a.date));
      if (s === 'lama') return String(a.date).localeCompare(String(b.date));
      return a.name.localeCompare(b.name, 'id');
    });
    return rows;
  }

  function renderAset() {
    var rows = filteredAssets();
    var total = rows.length;
    var pages = Math.max(1, Math.ceil(total / ui.pageSize));
    if (ui.page > pages) ui.page = pages;
    if (ui.page < 1) ui.page = 1;
    var start = (ui.page - 1) * ui.pageSize;
    var end = Math.min(start + ui.pageSize, total);
    var pageRows = rows.slice(start, end);
    var totalVal = rows.reduce(function (s, a) { return s + a.price; }, 0);

    var tbody = $('#aset-tbody');
    if (!pageRows.length) {
      tbody.innerHTML = '<tr><td colspan="7"><div class="empty-inline"><strong>Tidak ada aset yang cocok</strong><span>Coba ubah kata kunci atau filter yang dipilih.</span></div></td></tr>';
    } else {
      tbody.innerHTML = pageRows.map(function (a) {
        var sub = subName(a.sub);
        var loc = curLoc(a);
        var moved = state.overrides[a.ref] !== undefined;
        return '<tr data-ref="' + esc(a.ref) + '">' +
          '<td><span class="cell-kode">' + esc(a.ref) + '</span></td>' +
          '<td><div class="cell-name"><strong title="' + esc(a.name) + '">' + esc(a.name) + '</strong><small>' + esc(brandName(a) || a.desc || '—') + '</small></div></td>' +
          '<td><span class="badge ' + (SUB_BADGE[sub] || 'badge-gray') + '">' + esc(titleCase(sub)) + '</span></td>' +
          '<td><div class="cell-loc"><strong title="' + esc(locRaw(loc)[1]) + '">' + esc(locName(loc)) + (moved ? ' <span class="demo-tag" title="Lokasi sudah diperbarui dari riwayat">pindah</span>' : '') + '</strong><small>' + esc(locCode(loc) || '—') + '</small></div></td>' +
          '<td><span class="cell-date">' + fmtDate(a.date) + '</span></td>' +
          '<td class="num"><span class="cell-price">' + rp(a.price) + '</span></td>' +
          '<td><div class="row-actions">' +
            '<button type="button" data-detail="' + esc(a.ref) + '" aria-label="Lihat detail ' + esc(a.name) + '"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7"><path d="M2.5 12S6 5.5 12 5.5 21.5 12 21.5 12 18 18.5 12 18.5 2.5 12 2.5 12Z"/><circle cx="12" cy="12" r="3"/></svg></button>' +
            '<button type="button" data-move="' + esc(a.ref) + '" aria-label="Pindahkan ' + esc(a.name) + '"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7"><path d="M4 7h13l-3-3M20 17H7l3 3"/><path d="M17 7v4M7 17v-4"/></svg></button>' +
          '</div></td></tr>';
      }).join('');
    }

    $('#aset-result-text').innerHTML = total
      ? 'Menampilkan <strong>' + nf.format(start + 1) + '–' + nf.format(end) + '</strong> dari <strong>' + nf.format(total) + '</strong> aset · Total <strong>' + rp(totalVal) + '</strong>'
      : 'Tidak ada aset yang cocok dengan filter.';

    var chips = [];
    if (ui.search) chips.push(chip('Cari: ' + ui.search, 'search'));
    if (ui.sub !== '') chips.push(chip(titleCase(SUB[+ui.sub]), 'sub'));
    if (ui.loc !== '') chips.push(chip(locName(+ui.loc), 'loc'));
    $('#aset-active-filters').innerHTML = chips.join('');

    renderPagination(pages, total);

    $('#aset-pagesize').value = String(ui.pageSize);
  }

  function chip(text, type) {
    return '<span class="chip">' + esc(text) +
      '<button type="button" data-clear="' + type + '" aria-label="Hapus filter"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="m6 6 12 12M18 6 6 18"/></svg></button></span>';
  }

  function renderPagination(pages, total) {
    var wrap = $('#aset-pagination');
    if (total === 0) { wrap.innerHTML = ''; return; }
    var cur = ui.page;
    var items = [];
    var set = {};
    [1, pages, cur - 1, cur, cur + 1].forEach(function (p) { if (p >= 1 && p <= pages) set[p] = 1; });
    if (cur <= 3) [2, 3, 4].forEach(function (p) { if (p <= pages) set[p] = 1; });
    if (cur >= pages - 2) [pages - 1, pages - 2, pages - 3].forEach(function (p) { if (p >= 1) set[p] = 1; });
    var arr = Object.keys(set).map(Number).sort(function (a, b) { return a - b; });
    var prev = null;
    arr.forEach(function (p) {
      if (prev !== null && p - prev > 1) items.push('<span class="gap">…</span>');
      items.push('<button type="button" data-page="' + p + '" class="' + (p === cur ? 'is-active' : '') + '">' + p + '</button>');
      prev = p;
    });
    wrap.innerHTML =
      '<button type="button" data-page="' + (cur - 1) + '" ' + (cur === 1 ? 'disabled' : '') + ' aria-label="Sebelumnya"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="m14 6-6 6 6 6"/></svg></button>' +
      items.join('') +
      '<button type="button" data-page="' + (cur + 1) + '" ' + (cur === pages ? 'disabled' : '') + ' aria-label="Berikutnya"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="m10 6 6 6-6 6"/></svg></button>';
  }

  /* ---------------- Modal ---------------- */
  function openModal(inner) {
    var root = $('#modal-root');
    root.innerHTML = '<div class="modal-overlay" data-backdrop><div class="modal" role="dialog" aria-modal="true">' + inner + '</div></div>';
    document.body.style.overflow = 'hidden';
    var focusable = root.querySelector('button, [href], input, select, textarea');
    if (focusable) setTimeout(function () { focusable.focus(); }, 30);
  }
  function closeModal() {
    $('#modal-root').innerHTML = '';
    document.body.style.overflow = '';
  }

  function assetHistory(ref) {
    return sortedTransfers().filter(function (t) { return t.ref === ref; });
  }

  function openAssetDetail(ref) {
    var a = ASET.find(function (x) { return x.ref === ref; });
    if (!a) return;
    var loc = curLoc(a);
    var sub = subName(a.sub);
    var hist = assetHistory(a.ref);

    var html = '<div class="modal-head">' +
      '<div><h3>Detail Aset</h3><p>Informasi barang dan riwayat perpindahan</p></div>' +
      '<button class="icon-btn modal-close" data-close aria-label="Tutup"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="m6 6 12 12M18 6 6 18"/></svg></button>' +
      '</div><div class="modal-body">' +
      '<div class="detail-hero"><div class="dh-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="m12 3 8 4.5v9L12 21l-8-4.5v-9L12 3Z"/><path d="m4 7.5 8 4.5 8-4.5M12 12v9"/></svg></div>' +
      '<div><h4>' + esc(a.name) + '</h4><p>Kode aset ' + esc(a.ref) + ' · ' + esc(titleCase(sub)) + '</p></div></div>' +
      '<div class="detail-grid">' +
        detailItem('Kode Aset', a.ref, true) +
        detailItem('Tanggal Perolehan', fmtDate(a.date)) +
        detailItem('Nilai Perolehan', rp(a.price)) +
        detailItem('Merek / Detail', brandName(a) || '—') +
        detailItem('Sub Kelompok', titleCase(sub)) +
        detailItem('Kelompok', titleCase(KEL[a.kel] || '—')) +
        detailItem('Lokasi Saat Ini', locFull(loc), false, true) +
        detailItem('Deskripsi', a.desc || '—', false, true) +
      '</div>' +
      (hist.length ? '<p class="section-label">Riwayat perpindahan aset ini</p><div class="timeline">' + hist.map(function (t) {
        return '<div class="tl-item"><span class="tl-dot"></span><div class="tl-body">' +
          '<strong>' + esc(locName(t.from)) + ' → ' + esc(locName(t.to)) + '</strong>' +
          '<p>' + esc(t.reason) + ' · ' + esc(t.pic) + '</p>' +
          '<time>' + fmtDate(t.date) + (t.note ? ' · ' + esc(t.note) : '') + '</time>' +
          '</div></div>';
      }).join('') + '</div>' : '') +
      '</div><div class="modal-foot"><button class="btn btn-ghost" data-close> Tutup</button>' +
      '<button class="btn btn-primary" data-move="' + esc(a.ref) + '"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M4 7h13l-3-3M20 17H7l3 3"/><path d="M17 7v4M7 17v-4"/></svg> Pindahkan Aset</button></div>';

    openModal(html);
  }

  function detailItem(label, value, mono, full) {
    return '<div class="detail-item' + (full ? ' full' : '') + '"><small>' + esc(label) + '</small><span class="' + (mono ? 'mono' : '') + '">' + esc(value) + '</span></div>';
  }

  /* ---------------- Perpindahan ---------------- */
  function pickerSearch(q) {
    q = q.trim().toLowerCase();
    if (q.length < 2) return [];
    var out = [];
    for (var i = 0; i < ASET.length && out.length < 8; i++) {
      var a = ASET[i];
      var hay = (a.name + ' ' + a.ref + ' ' + locRaw(curLoc(a))[1]).toLowerCase();
      if (hay.indexOf(q) !== -1) out.push(a);
    }
    return out;
  }

  function renderPickerResults(q) {
    var box = $('#picker-results');
    var input = $('#asset-picker');
    var list = pickerSearch(q);
    if (q.trim().length < 2) {
      box.innerHTML = '<div class="picker-empty">Ketik minimal 2 karakter untuk mencari aset.</div>';
      box.hidden = false;
      input.setAttribute('aria-expanded', 'true');
      return;
    }
    if (!list.length) {
      box.innerHTML = '<div class="picker-empty">Tidak ada aset yang cocok.</div>';
      box.hidden = false;
      input.setAttribute('aria-expanded', 'true');
      return;
    }
    box.innerHTML = list.map(function (a) {
      return '<button type="button" class="picker-option" role="option" data-pick="' + esc(a.ref) + '">' +
        '<span class="opt-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="m12 3 8 4.5v9L12 21l-8-4.5v-9L12 3Z"/><path d="m4 7.5 8 4.5 8-4.5M12 12v9"/></svg></span>' +
        '<span class="opt-body"><strong>' + esc(a.name) + '</strong><small>' + esc(a.ref) + ' · ' + esc(locName(curLoc(a))) + '</small></span>' +
        '<span class="opt-price">' + rp(a.price) + '</span></button>';
    }).join('');
    box.hidden = false;
    input.setAttribute('aria-expanded', 'true');
  }

  function selectAsset(a) {
    ui.current = a;
    $('#asset-picker').value = a.name;
    $('#picker-results').hidden = true;
    $('#asset-picker').setAttribute('aria-expanded', 'false');
    $('#picker-clear').hidden = false;
    $('#move-from').value = locFull(curLoc(a));
    $('#move-from').classList.remove('is-error');

    var price = moneyShort(a.price);
    $('#selected-asset').hidden = false;
    $('#selected-asset').innerHTML =
      '<div class="sa-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="m12 3 8 4.5v9L12 21l-8-4.5v-9L12 3Z"/><path d="m4 7.5 8 4.5 8-4.5M12 12v9"/></svg></div>' +
      '<div class="sa-body"><strong>' + esc(a.name) + '</strong><span>Kode ' + esc(a.ref) + ' · ' + esc(titleCase(subName(a.sub))) + '</span></div>' +
      '<div class="sa-price">Rp ' + price.v + (price.u ? ' ' + price.u : '') + '</div>';

    updatePreview();
  }

  function clearAsset() {
    ui.current = null;
    $('#asset-picker').value = '';
    $('#asset-picker').setAttribute('aria-expanded', 'false');
    $('#picker-clear').hidden = true;
    $('#picker-results').hidden = true;
    $('#selected-asset').hidden = true;
    $('#move-from').value = '';
    updatePreview();
  }

  function updatePreview() {
    var a = ui.current;
    var body = $('#preview-body');
    var empty = $('#preview-empty');
    if (!a) {
      body.hidden = true; empty.hidden = false;
      return;
    }
    empty.hidden = true; body.hidden = false;

    var toVal = $('#move-to').value;
    var toId = toVal === '' ? null : +toVal;
    var date = $('#move-date').value;
    var pic = $('#move-pic').value;
    var reason = $('#move-reason').value;

    body.innerHTML =
      '<div class="preview-asset"><span class="pa-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="m12 3 8 4.5v9L12 21l-8-4.5v-9L12 3Z"/><path d="m4 7.5 8 4.5 8-4.5M12 12v9"/></svg></span>' +
      '<div><strong>' + esc(a.name) + '</strong><span>' + esc(a.ref) + ' · ' + rp(a.price) + '</span></div></div>' +
      '<div class="route">' +
        '<div class="route-point"><small>Dari</small><strong>' + esc(locName(curLoc(a))) + '</strong></div>' +
        '<div class="route-arrow"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M4 12h16m0 0-5-5m5 5-5 5"/></svg></div>' +
        '<div class="route-point"><small>Ke</small><strong>' + (toId === null ? '<span style="color:#9ab0be;font-weight:500">Belum dipilih</span>' : esc(locName(toId))) + '</strong></div>' +
      '</div>' +
      '<div class="preview-meta">' +
        '<div><span>Tanggal</span><strong>' + (date ? fmtDate(date) : '—') + '</strong></div>' +
        '<div><span>Petugas</span><strong>' + (pic ? esc(pic) : '—') + '</strong></div>' +
        '<div><span>Alasan</span><strong>' + (reason ? esc(reason) : '—') + '</strong></div>' +
        '<div><span>Lokasi kode</span><strong>' + (toId === null ? '—' : esc(locCode(toId) || '—')) + '</strong></div>' +
      '</div>';
  }

  function submitMove(e) {
    e.preventDefault();
    var a = ui.current;
    var toVal = $('#move-to').value;
    var date = $('#move-date').value;
    var pic = $('#move-pic').value.trim();
    var reason = $('#move-reason').value;
    var note = $('#move-note').value.trim();
    var err = $('#move-error');
    $$('#move-form .input, #move-form .picker input').forEach(function (el) { el.classList.remove('is-error'); });

    var msg = '';
    if (!a) msg = 'Pilih aset yang akan dipindahkan terlebih dahulu.';
    else if (toVal === '') msg = 'Lokasi tujuan wajib dipilih.';
    else if (+toVal === curLoc(a)) msg = 'Lokasi tujuan tidak boleh sama dengan lokasi asal.';
    else if (!date) msg = 'Tanggal perpindahan wajib diisi.';
    else if (!pic) msg = 'Nama petugas / PIC wajib diisi.';
    else if (!reason) msg = 'Alasan perpindahan wajib dipilih.';

    if (msg) {
      err.hidden = false;
      err.textContent = msg;
      if (!a) $('#asset-picker').classList.add('is-error');
      else if (toVal === '') $('#move-to').classList.add('is-error');
      else if (!date) $('#move-date').classList.add('is-error');
      else if (!pic) $('#move-pic').classList.add('is-error');
      else if (!reason) $('#move-reason').classList.add('is-error');
      return;
    }
    err.hidden = true;

    var id = 'MUT-' + Date.now();
    var from = curLoc(a);
    var transfer = {
      id: id,
      ref: a.ref,
      name: a.name,
      from: from,
      to: +toVal,
      date: date,
      pic: pic,
      reason: reason,
      note: note,
      createdAt: new Date().toISOString()
    };
    state.transfers.push(transfer);
    state.overrides[a.ref] = +toVal;
    save();

    var seq = state.transfers.length;
    var now = new Date();
    var noBukti = 'MUT/' + now.getFullYear() + '/' + String(now.getMonth() + 1).padStart(2, '0') + '/' + String(seq).padStart(3, '0');

    closeModal();
    openModal(
      '<div class="modal-head"><div><h3>Perpindahan Tersimpan</h3><p>Lokasi aset telah diperbarui</p></div>' +
      '<button class="icon-btn modal-close" data-close aria-label="Tutup"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="m6 6 12 12M18 6 6 18"/></svg></button></div>' +
      '<div class="modal-body"><div class="success-hero"><div class="success-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M5 12.5 10 17l9-10"/></svg></div>' +
      '<h3>Berhasil dicatat</h3><p>' + esc(a.name) + ' dipindahkan ke ' + esc(locName(+toVal)) + '.</p></div>' +
      '<div class="receipt">' +
        receiptRow('No. Bukti', noBukti) +
        receiptRow('Kode Aset', a.ref) +
        receiptRow('Dari', locFull(from)) +
        receiptRow('Ke', locFull(+toVal)) +
        receiptRow('Tanggal', fmtDate(date)) +
        receiptRow('Petugas', pic) +
        receiptRow('Alasan', reason) +
        (note ? receiptRow('Catatan', note) : '') +
      '</div></div>' +
      '<div class="modal-foot"><button class="btn btn-ghost" data-close>Selesai</button>' +
      '<button class="btn btn-primary" onclick="window.print()"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7"><path d="M7 8V4h10v4"/><rect x="4" y="8" width="16" height="8" rx="2"/><path d="M7 13h10v7H7z"/></svg> Cetak Bukti</button></div>'
    );

    resetMoveForm(true);
    toast('Perpindahan berhasil disimpan.');
    renderDashboard();
    if (ui.view === 'aset') renderAset();
    if (ui.view === 'riwayat') renderHistory();
  }

  function receiptRow(label, val) {
    return '<div class="receipt-row"><span>' + esc(label) + '</span><strong>' + esc(val) + '</strong></div>';
  }

  function resetMoveForm(keepDate) {
    clearAsset();
    $('#move-to').value = '';
    $('#move-pic').value = '';
    $('#move-reason').value = '';
    $('#move-note').value = '';
    if (!keepDate) $('#move-date').value = todayISO();
    $('#move-error').hidden = true;
    $$('#move-form .input, #move-form .picker input').forEach(function (el) { el.classList.remove('is-error'); });
    updatePreview();
  }

  /* ---------------- Riwayat ---------------- */
  function historyRows() {
    var q = ui.histSearch.trim().toLowerCase();
    return sortedTransfers().filter(function (t) {
      if (ui.histReason && t.reason !== ui.histReason) return false;
      if (!q) return true;
      var hay = (t.name + ' ' + t.ref + ' ' + t.pic + ' ' + t.reason + ' ' + t.note + ' ' +
        locRaw(t.from)[1] + ' ' + locRaw(t.to)[1]).toLowerCase();
      return hay.indexOf(q) !== -1;
    });
  }

  function renderHistory() {
    var all = state.transfers;
    var mm = all.filter(function (t) { return thisMonth(t.date); });
    var refs = {};
    all.forEach(function (t) { refs[t.ref] = 1; });

    $('#riwayat-stats').innerHTML =
      miniStat('Total perpindahan', nf.format(all.length), 'swap') +
      miniStat('Bulan ini', nf.format(mm.length), 'clock') +
      miniStat('Aset berbeda', nf.format(Object.keys(refs).length), 'box');

    var rows = historyRows();
    var tbody = $('#riwayat-tbody');
    $('#riwayat-empty').hidden = all.length > 0;
    $('#riwayat-table').closest('.table-scroll').style.display = all.length ? '' : 'none';

    if (!all.length) {
      tbody.innerHTML = '';
      $('#riwayat-result-text').textContent = '0 perpindahan';
      return;
    }

    $('#riwayat-result-text').innerHTML = rows.length
      ? 'Menampilkan <strong>' + nf.format(rows.length) + '</strong> perpindahan'
      : 'Tidak ada perpindahan yang cocok dengan filter.';

    if (!rows.length) {
      tbody.innerHTML = '<tr><td colspan="7"><div class="empty-inline"><strong>Tidak ada hasil</strong><span>Coba ubah kata kunci atau filter alasan.</span></div></td></tr>';
      return;
    }

    tbody.innerHTML = rows.map(function (t) {
      return '<tr>' +
        '<td><span class="cell-date">' + fmtDate(t.date) + '</span></td>' +
        '<td><div class="cell-name"><strong title="' + esc(t.name) + '">' + esc(t.name) + (t.demo ? '<span class="demo-tag">contoh</span>' : '') + '</strong><small>Kode ' + esc(t.ref) + '</small></div></td>' +
        '<td><div class="route-cell"><span class="lbl">Dari</span><span class="from">' + esc(locName(t.from)) + '</span></div></td>' +
        '<td><div class="route-cell"><span class="lbl">Ke</span><span class="to">' + esc(locName(t.to)) + '</span></div></td>' +
        '<td><span style="font-size:11.5px">' + esc(t.pic) + '</span></td>' +
        '<td><span class="badge ' + reasonBadge(t.reason) + '">' + esc(t.reason) + '</span></td>' +
        '<td><div class="row-actions">' +
          '<button type="button" data-detail="' + esc(t.ref) + '" aria-label="Detail aset"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7"><path d="M2.5 12S6 5.5 12 5.5 21.5 12 21.5 12 18 18.5 12 18.5 2.5 12 2.5 12Z"/><circle cx="12" cy="12" r="3"/></svg></button>' +
          '<button type="button" data-del="' + esc(t.id) + '" aria-label="Hapus catatan"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7"><path d="M4 7h16M9 7V5h6v2M6 7l1 13h10l1-13"/><path d="M10 11v5M14 11v5"/></svg></button>' +
        '</div></td></tr>';
    }).join('');
  }

  function miniStat(label, val, icon) {
    var paths = {
      swap: '<path d="M4 7h13l-3-3M20 17H7l3 3"/><path d="M17 7v4M7 17v-4"/>',
      clock: '<circle cx="12" cy="12" r="8.5"/><path d="M12 7.5v5l3 1.8"/>',
      box: '<path d="m12 3 8 4.5v9L12 21l-8-4.5v-9L12 3Z"/><path d="m4 7.5 8 4.5 8-4.5M12 12v9"/>'
    };
    return '<div class="mini-stat"><span class="ms-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round">' + paths[icon] + '</svg></span>' +
      '<div><strong>' + val + '</strong><span>' + esc(label) + '</span></div></div>';
  }

  function reasonBadge(r) {
    if (r === 'Relokasi ruang') return 'badge-blue';
    if (r === 'Kebutuhan pelayanan') return 'badge-teal';
    if (r === 'Perbaikan / servis') return 'badge-amber';
    if (r === 'Peminjaman unit lain') return 'badge-violet';
    if (r === 'Penggantian unit') return 'badge-pink';
    return 'badge-gray';
  }

  function confirmDelete(id) {
    var t = state.transfers.find(function (x) { return x.id === id; });
    if (!t) return;
    openModal(
      '<div class="modal-head"><div><h3>Hapus catatan perpindahan?</h3><p>Data aset akan dihitung ulang dari riwayat yang tersisa.</p></div>' +
      '<button class="icon-btn modal-close" data-close aria-label="Tutup"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="m6 6 12 12M18 6 6 18"/></svg></button></div>' +
      '<div class="modal-body"><div class="receipt">' +
      receiptRow('Aset', t.name) + receiptRow('Tanggal', fmtDate(t.date)) + receiptRow('Rute', locName(t.from) + ' → ' + locName(t.to)) + receiptRow('Petugas', t.pic) +
      '</div><p style="font-size:11.5px;color:var(--muted);margin-top:12px">Catatan ini akan dihapus permanen dari perangkat ini.</p></div>' +
      '<div class="modal-foot"><button class="btn btn-ghost" data-close>Batal</button>' +
      '<button class="btn btn-danger" data-confirm-del="' + esc(t.id) + '">Hapus catatan</button></div>'
    );
  }

  function doDelete(id) {
    state.transfers = state.transfers.filter(function (x) { return x.id !== id; });
    rebuildOverrides();
    save();
    closeModal();
    renderHistory();
    renderDashboard();
    toast('Catatan perpindahan dihapus.');
  }

  function confirmReset() {
    openModal(
      '<div class="modal-head"><div><h3>Reset seluruh data perpindahan?</h3><p>Termasuk 5 catatan contoh.</p></div>' +
      '<button class="icon-btn modal-close" data-close aria-label="Tutup"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="m6 6 12 12M18 6 6 18"/></svg></button></div>' +
      '<div class="modal-body"><p style="font-size:12px;color:var(--ink-soft);line-height:1.6">Semua riwayat perpindahan dan perubahan lokasi akan dihapus. Data aset dari Excel tetap utuh dan lokasi kembali seperti semula.</p></div>' +
      '<div class="modal-foot"><button class="btn btn-ghost" data-close>Batal</button>' +
      '<button class="btn btn-danger" data-confirm-reset>Ya, reset data</button></div>'
    );
  }

  /* ---------------- Dataset aktif ---------------- */
  function updateSourceCard() {
    $('#source-name').textContent = DATASET.source || 'Data Aset';
    $('#source-meta').textContent = nf.format(ASET.length) + ' unit · ' + nf.format(uniqueLocations()) + ' lokasi';
  }

  function pruneTransfers() {
    var refs = {};
    ASET.forEach(function (a) { refs[a.ref] = 1; });
    state.transfers = state.transfers.filter(function (t) { return refs[t.ref]; });
    rebuildOverrides();
    save();
  }

  function persistDataset() {
    if (!window.DatasetStore) return Promise.resolve(false);
    return window.DatasetStore.set({ db: DB, info: DATASET, savedAt: new Date().toISOString() })
      .then(function () { return true; })
      .catch(function () {
        toast('Data aktif berjalan, tetapi gagal disimpan permanen di peramban ini.');
        return false;
      });
  }

  function applyDataset(db, info) {
    buildData(db);
    DATASET = {
      kind: (info && info.kind) || 'custom',
      source: (info && info.source) || (db.meta && db.meta.source) || 'Data Aset',
      fileName: (info && info.fileName) || '',
      sheet: (info && info.sheet) || '',
      updatedAt: (info && info.updatedAt) || new Date().toISOString(),
      importedAt: (info && info.importedAt) || ''
    };

    pruneTransfers();
    buildFilters();

    ui.page = 1;
    ui.search = '';
    ui.sub = '';
    ui.loc = '';
    if ($('#aset-search')) $('#aset-search').value = '';
    if ($('#aset-sub')) $('#aset-sub').value = '';
    if ($('#aset-lokasi')) $('#aset-lokasi').value = '';
    ui.depSearch = '';
    ui.depKat = '';
    ui.depStatus = '';
    ui.depPage = 1;
    if ($('#dep-search')) $('#dep-search').value = '';
    if ($('#dep-kat')) $('#dep-kat').value = '';
    if ($('#dep-status')) $('#dep-status').value = '';
    ui.current = null;
    ui.repairAsset = null;
    clearAsset();
    clearRepairAsset();

    renderAll();
    updateSourceCard();
    renderAddPreview();
    renderRepairs();
    updatePreview();
  }

  /* ---------------- Tambah barang ---------------- */
  function resetAddForm() {
    ['add-ref', 'add-name', 'add-brand', 'add-desc', 'add-price'].forEach(function (id) {
      var el = $('#' + id);
      if (el) el.value = '';
    });
    if ($('#add-sub')) $('#add-sub').value = '';
    if ($('#add-loc')) $('#add-loc').value = '';
    if ($('#add-date')) $('#add-date').value = todayISO();
    if ($('#add-kel')) $('#add-kel').value = 'AKTIVA TETAP BERWUJUD';
    var err = $('#add-error');
    if (err) err.hidden = true;
    $$('#add-form .input, #add-form .select').forEach(function (el) { el.classList.remove('is-error'); });
    renderAddPreview();
  }

  function renderAddPreview() {
    var body = $('#add-preview-body');
    var empty = $('#add-preview-empty');
    if (!body || !empty) return;

    var ref = $('#add-ref').value.trim();
    var name = $('#add-name').value.trim();
    var price = parseMoneyValue($('#add-price').value);
    var subVal = $('#add-sub').value;
    var locVal = $('#add-loc').value;
    var date = $('#add-date').value;
    var brand = $('#add-brand').value.trim();
    var desc = $('#add-desc').value.trim();
    var kel = $('#add-kel').value.trim();

    if (!ref && !name && !price && subVal === '' && locVal === '') {
      body.hidden = true;
      empty.hidden = false;
      return;
    }
    empty.hidden = true;
    body.hidden = false;

    var estimate = Math.round(price * REPAIR_PERCENT);
    var dup = ref && ASET.some(function (a) { return a.ref.toLowerCase() === ref.toLowerCase(); });

    body.innerHTML =
      '<div class="preview-asset"><span class="pa-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="m12 3 8 4.5v9L12 21l-8-4.5v-9L12 3Z"/><path d="m4 7.5 8 4.5 8-4.5M12 12v9"/></svg></span>' +
      '<div><strong>' + esc(name || 'Barang baru') + '</strong><span>' + esc(ref ? 'Kode ' + ref : 'Kode belum diisi') + '</span></div></div>' +
      (dup ? '<p class="preview-warn"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="9"/><path d="M12 8v5M12 16h.01"/></svg> Kode aset sudah terdaftar</p>' : '') +
      '<div class="preview-meta">' +
        '<div><span>Sub kelompok</span><strong>' + (subVal === '' ? '—' : esc(titleCase(subName(+subVal)))) + '</strong></div>' +
        '<div><span>Lokasi</span><strong>' + (locVal === '' ? '—' : esc(locName(+locVal))) + '</strong></div>' +
        '<div><span>Tanggal perolehan</span><strong>' + (date ? fmtDate(date) : '—') + '</strong></div>' +
        '<div><span>Harga perolehan</span><strong>' + (price ? rp(price) : '—') + '</strong></div>' +
        '<div><span>Estimasi 15%</span><strong>' + (price ? rp(estimate) : '—') + '</strong></div>' +
        '<div><span>Merek / detail</span><strong>' + (brand ? esc(brand) : '—') + '</strong></div>' +
      '</div>' +
      (desc ? '<p class="preview-desc">' + esc(desc) + '</p>' : '') +
      (kel ? '<p class="preview-desc"><span class="kv">Kelompok</span> ' + esc(titleCase(kel)) + '</p>' : '');
  }

  function submitAdd(e) {
    e.preventDefault();
    var ref = $('#add-ref').value.trim();
    var name = $('#add-name').value.trim();
    var subVal = $('#add-sub').value;
    var locVal = $('#add-loc').value;
    var date = $('#add-date').value;
    var price = parseMoneyValue($('#add-price').value);
    var brandTxt = $('#add-brand').value.trim();
    var kelTxt = $('#add-kel').value.trim() || 'AKTIVA TETAP BERWUJUD';
    var desc = $('#add-desc').value.trim();
    var err = $('#add-error');

    $$('#add-form .input, #add-form .select').forEach(function (el) { el.classList.remove('is-error'); });
    var msg = '';
    if (!ref) msg = 'Kode aset wajib diisi.';
    else if (ASET.some(function (a) { return a.ref.toLowerCase() === ref.toLowerCase(); })) msg = 'Kode aset "' + ref + '" sudah terdaftar. Gunakan kode lain.';
    else if (!name) msg = 'Nama barang wajib diisi.';
    else if (subVal === '') msg = 'Sub kelompok wajib dipilih.';
    else if (locVal === '') msg = 'Lokasi wajib dipilih.';
    else if (!date) msg = 'Tanggal perolehan wajib diisi.';
    else if (price <= 0) msg = 'Harga perolehan wajib diisi dan lebih dari 0.';

    if (msg) {
      err.hidden = false;
      err.textContent = msg;
      if (!ref || ASET.some(function (a) { return a.ref.toLowerCase() === ref.toLowerCase(); })) $('#add-ref').classList.add('is-error');
      else if (!name) $('#add-name').classList.add('is-error');
      else if (subVal === '') $('#add-sub').classList.add('is-error');
      else if (locVal === '') $('#add-loc').classList.add('is-error');
      else if (!date) $('#add-date').classList.add('is-error');
      else if (price <= 0) $('#add-price').classList.add('is-error');
      return;
    }
    err.hidden = true;

    var base = cloneDb(DB);
    var brandId = -1;
    if (brandTxt) {
      var found = -1;
      base.merek.forEach(function (m, i) { if (String(m).toLowerCase() === brandTxt.toLowerCase()) found = i; });
      if (found === -1) { base.merek.push(brandTxt); found = base.merek.length - 1; }
      brandId = found;
    }
    var kelId = -1;
    base.kel.forEach(function (k, i) { if (String(k).toUpperCase() === kelTxt.toUpperCase()) kelId = i; });
    if (kelId === -1) { base.kel.push(kelTxt); kelId = base.kel.length - 1; }

    var raw = [ref, name, +locVal, date, +subVal, brandId, price, kelId, desc];
    base.aset = base.aset.concat([raw]);

    var info = {
      kind: 'custom',
      source: DATASET.source || 'Data Aset',
      fileName: DATASET.fileName || '',
      sheet: DATASET.sheet || '',
      updatedAt: new Date().toISOString()
    };
    withMeta(base, info);
    applyDataset(base, info);
    persistDataset();
    addLog('Tambah barang', ref + ' · ' + name + ' · ' + rp(price));

    openModal(
      '<div class="modal-head"><div><h3>Barang Tersimpan</h3><p>Aset baru berhasil ditambahkan</p></div>' +
      '<button class="icon-btn modal-close" data-close aria-label="Tutup"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="m6 6 12 12M18 6 6 18"/></svg></button></div>' +
      '<div class="modal-body"><div class="success-hero"><div class="success-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M5 12.5 10 17l9-10"/></svg></div>' +
      '<h3>Berhasil ditambahkan</h3><p>' + esc(name) + ' sudah masuk ke data aset.</p></div>' +
      '<div class="receipt">' +
        receiptRow('Kode Aset', ref) +
        receiptRow('Nama Barang', name) +
        receiptRow('Sub Kelompok', titleCase(subName(+subVal))) +
        receiptRow('Lokasi', locFull(+locVal)) +
        receiptRow('Tanggal Perolehan', fmtDate(date)) +
        receiptRow('Harga Perolehan', rp(price)) +
        receiptRow('Estimasi 15%', rp(Math.round(price * REPAIR_PERCENT))) +
      '</div></div>' +
      '<div class="modal-foot"><button class="btn btn-ghost" data-goto="aset">Lihat Data Aset</button>' +
      '<button class="btn btn-primary" data-close>Tambah lagi</button></div>'
    );

    resetAddForm();
    toast('Barang baru berhasil ditambahkan.');
  }

  /* ---------------- Perbaikan ---------------- */
  function repairCalc() {
    var a = ui.repairAsset;
    var price = a ? a.price : 0;
    var estimate = Math.round(price * REPAIR_PERCENT);
    var cost = parseMoneyValue($('#repair-cost') ? $('#repair-cost').value : '');
    return {
      asset: a, price: price, estimate: estimate, cost: cost,
      okCost: cost > REPAIR_THRESHOLD,
      okEstimate: estimate > REPAIR_THRESHOLD,
      eligible: cost > REPAIR_THRESHOLD && estimate > REPAIR_THRESHOLD
    };
  }

  function renderRepair() {
    if (!$('#repair-price')) return;
    var a = ui.repairAsset;
    $('#repair-price').value = a ? rp(a.price) : '';
    $('#repair-estimate').value = a ? rp(Math.round(a.price * REPAIR_PERCENT)) : '';
    renderRepairResult();
  }

  function renderRepairResult() {
    var body = $('#repair-result-body');
    if (!body) return;
    var r = repairCalc();
    if (!r.asset) {
      body.innerHTML = '<div class="preview-empty"><div class="preview-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M9 11a3 3 0 1 0 6 0 3 3 0 1 0-6 0"/><path d="M12 14v7"/><path d="M8 21h8"/></svg></div><p>Pilih aset dan isi biaya perbaikan untuk melihat hasil.</p></div>';
      return;
    }
    var waiting = !r.cost;
    var verdictClass = waiting ? 'wait' : (r.eligible ? 'pass' : 'fail');
    var verdictIcon = waiting
      ? '<path d="M12 8v5M12 16h.01"/><circle cx="12" cy="12" r="9"/>'
      : (r.eligible ? '<path d="M5 12.5 10 17l9-10"/>' : '<path d="m7 7 10 10M17 7 7 17"/>');
    var verdictTitle = waiting ? 'Menunggu biaya perbaikan' : (r.eligible ? 'Memenuhi syarat perbaikan' : 'Tidak memenuhi syarat perbaikan');
    var verdictText = waiting
      ? 'Masukkan biaya perbaikan untuk memeriksa kelayakan.'
      : (r.eligible
        ? 'Kedua syarat terpenuhi: biaya perbaikan dan estimasi 15% di atas Rp 4.000.000.'
        : 'Salah satu atau kedua syarat belum terpenuhi.');

    body.innerHTML =
      '<div class="verdict verdict-' + verdictClass + '">' +
        '<div class="verdict-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">' + verdictIcon + '</svg></div>' +
        '<strong>' + verdictTitle + '</strong><p>' + verdictText + '</p></div>' +
      '<div class="calc-list">' +
        '<div class="calc-row"><span>Harga perolehan</span><strong>' + rp(r.price) + '</strong></div>' +
        '<div class="calc-row"><span>Estimasi 15%</span><strong>' + rp(r.estimate) + '</strong></div>' +
        '<div class="calc-row"><span>Biaya perbaikan</span><strong>' + (r.cost ? rp(r.cost) : '—') + '</strong></div>' +
      '</div>' +
      '<div class="check-list">' +
        checkRow('Biaya perbaikan > Rp 4.000.000', r.cost ? rp(r.cost) : 'Belum diisi', waiting ? null : r.okCost) +
        checkRow('Estimasi 15% > Rp 4.000.000', rp(r.estimate), r.okEstimate) +
      '</div>';
  }

  function checkRow(label, value, pass) {
    var cls = pass === null ? 'is-wait' : (pass ? 'is-pass' : 'is-fail');
    var icon = pass === null
      ? '<circle cx="12" cy="12" r="9"/><path d="M12 8v5M12 16h.01"/>'
      : (pass ? '<path d="M5 12.5 10 17l9-10"/>' : '<path d="m7 7 10 10M17 7 7 17"/>');
    return '<div class="check-row ' + cls + '"><span class="check-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">' + icon + '</svg></span>' +
      '<div><strong>' + esc(label) + '</strong><small>' + esc(value) + '</small></div></div>';
  }

  function repairPickerSearch(q) {
    q = q.trim().toLowerCase();
    if (q.length < 2) return [];
    var out = [];
    for (var i = 0; i < ASET.length && out.length < 8; i++) {
      var a = ASET[i];
      var hay = (a.name + ' ' + a.ref + ' ' + locRaw(curLoc(a))[1]).toLowerCase();
      if (hay.indexOf(q) !== -1) out.push(a);
    }
    return out;
  }

  function renderRepairPickerResults(q) {
    var box = $('#repair-picker-results');
    var input = $('#repair-picker');
    var list = repairPickerSearch(q);
    if (q.trim().length < 2) {
      box.innerHTML = '<div class="picker-empty">Ketik minimal 2 karakter untuk mencari aset.</div>';
    } else if (!list.length) {
      box.innerHTML = '<div class="picker-empty">Tidak ada aset yang cocok.</div>';
    } else {
      box.innerHTML = list.map(function (a) {
        return '<button type="button" class="picker-option" role="option" data-repair-pick="' + esc(a.ref) + '">' +
          '<span class="opt-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="m12 3 8 4.5v9L12 21l-8-4.5v-9L12 3Z"/><path d="m4 7.5 8 4.5 8-4.5M12 12v9"/></svg></span>' +
          '<span class="opt-body"><strong>' + esc(a.name) + '</strong><small>' + esc(a.ref) + ' · ' + esc(locName(curLoc(a))) + '</small></span>' +
          '<span class="opt-price">' + rp(a.price) + '</span></button>';
      }).join('');
    }
    box.hidden = false;
    input.setAttribute('aria-expanded', 'true');
  }

  function selectRepairAsset(a) {
    ui.repairAsset = a;
    $('#repair-picker').value = a.name;
    $('#repair-picker-results').hidden = true;
    $('#repair-picker').setAttribute('aria-expanded', 'false');
    $('#repair-picker-clear').hidden = false;
    $('#repair-price').value = rp(a.price);
    $('#repair-estimate').value = rp(Math.round(a.price * REPAIR_PERCENT));

    $('#repair-selected').hidden = false;
    $('#repair-selected').innerHTML =
      '<div class="sa-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="m12 3 8 4.5v9L12 21l-8-4.5v-9L12 3Z"/><path d="m4 7.5 8 4.5 8-4.5M12 12v9"/></svg></div>' +
      '<div class="sa-body"><strong>' + esc(a.name) + '</strong><span>Kode ' + esc(a.ref) + ' · ' + esc(locName(curLoc(a))) + '</span></div>' +
      '<div class="sa-price">' + rp(a.price) + '</div>';
    renderRepairResult();
  }

  function clearRepairAsset() {
    ui.repairAsset = null;
    if (!$('#repair-picker')) return;
    $('#repair-picker').value = '';
    $('#repair-picker').setAttribute('aria-expanded', 'false');
    $('#repair-picker-clear').hidden = true;
    $('#repair-picker-results').hidden = true;
    $('#repair-selected').hidden = true;
    $('#repair-price').value = '';
    $('#repair-estimate').value = '';
    renderRepairResult();
  }

  function resetRepairForm() {
    clearRepairAsset();
    if ($('#repair-cost')) $('#repair-cost').value = '';
    if ($('#repair-error')) $('#repair-error').hidden = true;
    $$('#repair-form .input').forEach(function (el) { el.classList.remove('is-error'); });
    renderRepairResult();
  }

  function submitRepair(e) {
    e.preventDefault();
    var r = repairCalc();
    var err = $('#repair-error');
    $$('#repair-form .input').forEach(function (el) { el.classList.remove('is-error'); });
    var msg = '';
    if (!r.asset) msg = 'Pilih aset terlebih dahulu.';
    else if (!r.cost) msg = 'Biaya perbaikan wajib diisi dan lebih dari 0.';
    if (msg) {
      err.hidden = false;
      err.textContent = msg;
      if (!r.asset) $('#repair-picker').classList.add('is-error');
      else $('#repair-cost').classList.add('is-error');
      return;
    }
    err.hidden = true;

    repairs.unshift({
      id: 'RPR-' + Date.now(),
      ref: r.asset.ref,
      name: r.asset.name,
      price: r.price,
      estimate: r.estimate,
      cost: r.cost,
      eligible: r.eligible,
      date: todayISO(),
      createdAt: new Date().toISOString()
    });
    saveRepairs();
    renderRepairs();
    addLog('Hitung perbaikan', r.asset.name + ' · biaya ' + rp(r.cost) + ' · ' + (r.eligible ? 'memenuhi syarat' : 'tidak memenuhi'));
    toast(r.eligible ? 'Hasil disimpan: memenuhi syarat perbaikan.' : 'Hasil disimpan: tidak memenuhi syarat perbaikan.');
  }

  function loadRepairs() {
    try {
      var raw = localStorage.getItem(REPAIR_KEY);
      var parsed = raw ? JSON.parse(raw) : null;
      if (Array.isArray(parsed)) repairs = parsed;
    } catch (e) { repairs = []; }
  }

  function saveRepairs() {
    try { localStorage.setItem(REPAIR_KEY, JSON.stringify(repairs)); } catch (e) { /* mode privat */ }
  }

  function renderRepairs() {
    var tbody = $('#repair-tbody');
    if (!tbody) return;
    var wrap = $('#repair-table-wrap');
    var empty = $('#repair-empty');
    var total = repairs.length;
    var pass = repairs.filter(function (r) { return r.eligible; }).length;
    var text = $('#repair-result-text');
    if (text) {
      text.innerHTML = total
        ? 'Menampilkan <strong>' + nf.format(total) + '</strong> perhitungan · <strong>' + nf.format(pass) + '</strong> memenuhi syarat'
        : '0 perhitungan';
    }
    if (empty) empty.hidden = total > 0;
    if (wrap) wrap.style.display = total ? '' : 'none';

    if (!total) { tbody.innerHTML = ''; return; }
    tbody.innerHTML = repairs.map(function (r) {
      return '<tr>' +
        '<td><span class="cell-date">' + fmtDate(r.date) + '</span></td>' +
        '<td><div class="cell-name"><strong title="' + esc(r.name) + '">' + esc(r.name) + '</strong><small>Kode ' + esc(r.ref) + '</small></div></td>' +
        '<td class="num"><span class="cell-price">' + rp(r.price) + '</span></td>' +
        '<td class="num"><span class="cell-price">' + rp(r.estimate) + '</span></td>' +
        '<td class="num"><span class="cell-price">' + rp(r.cost) + '</span></td>' +
        '<td><span class="badge ' + (r.eligible ? 'badge-teal' : 'badge-pink') + '">' + (r.eligible ? 'Memenuhi syarat' : 'Tidak memenuhi') + '</span></td>' +
        '<td><div class="row-actions">' +
          '<button type="button" data-repair-load="' + esc(r.id) + '" aria-label="Hitung lagi"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7"><path d="M4 4v6h6"/><path d="M20 20v-6h-6"/><path d="M6.5 16.5a8 8 0 0 0 13-3M17.5 7.5a8 8 0 0 0-13 3"/></svg></button>' +
          '<button type="button" data-repair-del="' + esc(r.id) + '" aria-label="Hapus hasil"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7"><path d="M4 7h16M9 7V5h6v2M6 7l1 13h10l1-13"/><path d="M10 11v5M14 11v5"/></svg></button>' +
        '</div></td></tr>';
    }).join('');
  }

  function loadRepairRecord(id) {
    var r = repairs.find(function (x) { return x.id === id; });
    if (!r) return;
    var a = ASET.find(function (x) { return x.ref === r.ref; });
    if (!a) { toast('Aset untuk perhitungan ini sudah tidak ada di dataset.'); return; }
    go('perbaikan');
    selectRepairAsset(a);
    $('#repair-cost').value = nf.format(r.cost);
    renderRepairResult();
    setTimeout(function () { $('#repair-cost').focus(); }, 120);
  }

  function confirmDeleteRepair(id) {
    var r = repairs.find(function (x) { return x.id === id; });
    if (!r) return;
    openModal(
      '<div class="modal-head"><div><h3>Hapus hasil perhitungan?</h3><p>Catatan ini akan dihapus dari perangkat.</p></div>' +
      '<button class="icon-btn modal-close" data-close aria-label="Tutup"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="m6 6 12 12M18 6 6 18"/></svg></button></div>' +
      '<div class="modal-body"><div class="receipt">' + receiptRow('Aset', r.name) + receiptRow('Harga', rp(r.price)) + receiptRow('Estimasi 15%', rp(r.estimate)) + receiptRow('Biaya', rp(r.cost)) + '</div></div>' +
      '<div class="modal-foot"><button class="btn btn-ghost" data-close>Batal</button><button class="btn btn-danger" data-confirm-repair-del="' + esc(r.id) + '">Hapus hasil</button></div>'
    );
  }

  function doDeleteRepair(id) {
    repairs = repairs.filter(function (x) { return x.id !== id; });
    saveRepairs();
    renderRepairs();
    closeModal();
    toast('Hasil perhitungan dihapus.');
  }

  function exportRepairs() {
    var rows = [['Tanggal', 'Kode Aset', 'Nama Aset', 'Harga Perolehan', 'Estimasi 15%', 'Biaya Perbaikan', 'Status']];
    repairs.forEach(function (r) {
      rows.push([r.date, r.ref, r.name, r.price, r.estimate, r.cost, r.eligible ? 'Memenuhi syarat' : 'Tidak memenuhi syarat']);
    });
    downloadCSV('riwayat-perbaikan-' + todayISO() + '.csv', rows);
    toast(rows.length - 1 + ' perhitungan diekspor ke CSV.');
  }

  /* ---------------- Penyusutan (garis lurus) ---------------- */
  function depKatCounts() {
    var list = KAT_LIST.map(function () { return 0; });
    var unknown = 0;
    ASET.forEach(function (a) {
      var i = katIndexOf(a);
      if (i < 0) unknown++;
      else list[i]++;
    });
    return { list: list, unknown: unknown };
  }

  function monthsElapsed(iso) {
    var d = dateObj(iso);
    if (!d) return 0;
    var now = new Date();
    var m = (now.getFullYear() - d.getFullYear()) * 12 + (now.getMonth() - d.getMonth());
    if (now.getDate() < d.getDate()) m--;
    return Math.max(0, m);
  }

  function depOf(a) {
    var ki = katIndexOf(a);
    var years = ki >= 0 ? (+KAT_LIFE[ki] || 0) : 0;
    var annual = years > 0 ? a.price / years : 0;
    var monthly = annual / 12;
    var lifeMonths = years * 12;
    var ageMonths = monthsElapsed(a.date);
    var chargedMonths = Math.min(ageMonths, lifeMonths);
    var accumulated = years > 0 ? Math.min(a.price, Math.round(monthly * chargedMonths)) : 0;
    return {
      ki: ki,
      years: years,
      annual: annual,
      monthly: monthly,
      lifeMonths: lifeMonths,
      ageMonths: ageMonths,
      accumulated: accumulated,
      book: Math.max(0, a.price - accumulated),
      done: years > 0 && ageMonths >= lifeMonths
    };
  }

  function formatAge(months) {
    if (!months) return '0 bulan';
    var y = Math.floor(months / 12);
    var m = months % 12;
    if (!y) return m + ' bulan';
    return y + ' th' + (m ? ' ' + m + ' bln' : '');
  }

  function filteredDep() {
    var q = ui.depSearch.trim().toLowerCase();
    var kat = ui.depKat;
    var st = ui.depStatus;
    var rows = ASET.map(function (a) { return { a: a, d: depOf(a) }; }).filter(function (x) {
      if (kat === 'unknown') {
        if (x.d.ki !== -1) return false;
      } else if (kat !== '' && x.d.ki !== +kat) {
        return false;
      }
      if (st === 'active' && (!x.d.years || x.d.done)) return false;
      if (st === 'done' && !x.d.done) return false;
      if (!q) return true;
      var l = curLoc(x.a);
      var hay = (x.a.name + ' ' + x.a.ref + ' ' + katNameOf(x.a) + ' ' + brandName(x.a) + ' ' +
        locName(l) + ' ' + locCode(l)).toLowerCase();
      return hay.indexOf(q) !== -1;
    });

    var s = ui.depSort;
    rows.sort(function (x, y) {
      if (s === 'tahunan-desc') return y.d.annual - x.d.annual;
      if (s === 'nilai-desc') return y.a.price - x.a.price;
      if (s === 'baru') return String(y.a.date).localeCompare(String(x.a.date));
      if (s === 'lama') return String(x.a.date).localeCompare(String(y.a.date));
      return x.a.name.localeCompare(y.a.name, 'id');
    });
    return rows;
  }

  function depStatusBadge(d) {
    if (!d.years) return '<span class="badge badge-gray">Tanpa kategori</span>';
    if (d.done) return '<span class="badge badge-gray">Selesai</span>';
    return '<span class="badge badge-teal">Berjalan</span>';
  }

  function renderPenyusutan() {
    if (!$('#dep-tbody')) return;

    var all = ASET.map(function (a) { return { a: a, d: depOf(a) }; });
    var totalValue = 0, totalAnnual = 0, totalMonthly = 0, doneCount = 0;
    all.forEach(function (x) {
      totalValue += x.a.price;
      totalAnnual += x.d.annual;
      totalMonthly += x.d.monthly;
      if (x.d.done) doneCount++;
    });

    $('#penyusutan-kpis').innerHTML = [
      kpi('Jumlah Aset', nf.format(ASET.length), 'unit', 'layers', '#2f7de1', '#e9f1fd',
        KAT_LIST.length + ' kategori masa manfaat'),
      kpi('Total Harga Perolehan', 'Rp ' + moneyShort(totalValue).v, moneyShort(totalValue).u, 'wallet', '#0d9488', '#e3f6f2',
        'Nilai residu Rp 0'),
      kpi('Penyusutan Tahunan', 'Rp ' + moneyShort(totalAnnual).v, moneyShort(totalAnnual).u, 'trend', '#e59a17', '#fdf3de',
        'Harga perolehan ÷ masa manfaat'),
      kpi('Penyusutan Bulanan', 'Rp ' + moneyShort(totalMonthly).v, moneyShort(totalMonthly).u, 'calendar', '#8b5cf6', '#eeecfe',
        doneCount + ' aset sudah habis masa manfaat')
    ].join('');

    var counts = depKatCounts();
    var catsHtml = KAT_LIST.map(function (name, i) {
      var value = 0;
      ASET.forEach(function (a) { if (katIndexOf(a) === i) value += a.price; });
      return '<div class="dep-cat">' +
        '<span class="dep-dot" style="background:' + KAT_DOT[i] + '"></span>' +
        '<div class="dep-cat-main"><strong>' + esc(name) + '</strong>' +
        '<small>' + nf.format(counts.list[i]) + ' aset · ' + rupiahShortText(value) + '</small></div>' +
        '<span class="dep-cat-life">' + (+KAT_LIFE[i] || 0) + ' th</span></div>';
    });
    if (counts.unknown) {
      catsHtml.push('<div class="dep-cat dep-cat-unknown">' +
        '<span class="dep-dot" style="background:#94a9b8"></span>' +
        '<div class="dep-cat-main"><strong>Belum berkategori</strong>' +
        '<small>' + nf.format(counts.unknown) + ' aset · tanpa masa manfaat</small></div>' +
        '<span class="dep-cat-life">—</span></div>');
    }
    $('#dep-cats').innerHTML = catsHtml.join('');

    var rows = filteredDep();
    var total = rows.length;
    var pages = Math.max(1, Math.ceil(total / ui.depPageSize));
    if (ui.depPage > pages) ui.depPage = pages;
    if (ui.depPage < 1) ui.depPage = 1;
    var start = (ui.depPage - 1) * ui.depPageSize;
    var end = Math.min(start + ui.depPageSize, total);
    var pageRows = rows.slice(start, end);

    var annualSum = rows.reduce(function (s, x) { return s + x.d.annual; }, 0);
    var monthlySum = rows.reduce(function (s, x) { return s + x.d.monthly; }, 0);

    if (!pageRows.length) {
      $('#dep-tbody').innerHTML = '<tr><td colspan="9"><div class="empty-inline"><strong>Tidak ada aset yang cocok</strong><span>Coba ubah kata kunci, kategori, atau status penyusutan.</span></div></td></tr>';
    } else {
      $('#dep-tbody').innerHTML = pageRows.map(function (x) {
        var a = x.a, d = x.d, l = curLoc(a);
        return '<tr data-dep-ref="' + esc(a.ref) + '">' +
          '<td><div class="cell-name"><strong title="' + esc(a.name) + '">' + esc(a.name) + '</strong>' +
          '<small>' + esc(a.ref) + ' · ' + esc(locName(l)) + '</small></div></td>' +
          '<td>' + (d.ki >= 0
            ? '<span class="kat-badge ' + KAT_BADGE[d.ki] + '">' + esc(KAT_LIST[d.ki]) + '</span>'
            : '<span class="kat-badge kat-unknown">Belum berkategori</span>') + '</td>' +
          '<td><span class="cell-date">' + fmtDate(a.date) + '</span></td>' +
          '<td>' + (d.years
            ? '<span class="dep-life">' + d.years + ' tahun</span><small class="dep-life-sub">' + d.lifeMonths + ' bulan</small>'
            : '<span class="dep-life muted">—</span>') + '</td>' +
          '<td class="num"><span class="cell-price">' + rp(a.price) + '</span></td>' +
          '<td class="num"><strong class="dep-annual">' + (d.years ? rp(Math.round(d.annual)) : '—') + '</strong></td>' +
          '<td class="num">' + (d.years ? rp(Math.round(d.monthly)) : '—') + '</td>' +
          '<td>' + depStatusBadge(d) + '<small class="dep-age">' + (d.years ? formatAge(d.ageMonths) : '—') + '</small></td>' +
          '<td><div class="row-actions">' +
            '<button type="button" data-dep-detail="' + esc(a.ref) + '" aria-label="Detail penyusutan ' + esc(a.name) + '"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7"><path d="M2.5 12S6 5.5 12 5.5 21.5 12 21.5 12 18 18.5 12 18.5 2.5 12 2.5 12Z"/><circle cx="12" cy="12" r="3"/></svg></button>' +
          '</div></td></tr>';
      }).join('');
    }

    $('#dep-result-text').innerHTML = total
      ? 'Menampilkan <strong>' + nf.format(start + 1) + '–' + nf.format(end) + '</strong> dari <strong>' + nf.format(total) + '</strong> aset · Penyusutan tahunan <strong>' + rp(Math.round(annualSum)) + '</strong> · bulanan <strong>' + rp(Math.round(monthlySum)) + '</strong>'
      : 'Tidak ada aset yang cocok dengan filter.';

    var chips = [];
    if (ui.depSearch) chips.push(chip('Cari: ' + ui.depSearch, 'dep-search'));
    if (ui.depKat !== '') {
      chips.push(chip(ui.depKat === 'unknown' ? 'Belum berkategori' : KAT_LIST[+ui.depKat], 'dep-kat'));
    }
    if (ui.depStatus === 'active') chips.push(chip('Masih disusutkan', 'dep-status'));
    if (ui.depStatus === 'done') chips.push(chip('Masa manfaat selesai', 'dep-status'));
    $('#dep-active-filters').innerHTML = chips.join('');

    renderDepPagination(pages, total);

    $('#dep-pagesize').value = String(ui.depPageSize);
  }

  function renderDepPagination(pages, total) {
    var wrap = $('#dep-pagination');
    if (!wrap) return;
    if (total === 0) { wrap.innerHTML = ''; return; }
    var cur = ui.depPage;
    var items = [];
    var set = {};
    [1, pages, cur - 1, cur, cur + 1].forEach(function (p) { if (p >= 1 && p <= pages) set[p] = 1; });
    if (cur <= 3) [2, 3, 4].forEach(function (p) { if (p <= pages) set[p] = 1; });
    if (cur >= pages - 2) [pages - 1, pages - 2, pages - 3].forEach(function (p) { if (p >= 1) set[p] = 1; });
    var arr = Object.keys(set).map(Number).sort(function (a, b) { return a - b; });
    var prev = null;
    arr.forEach(function (p) {
      if (prev !== null && p - prev > 1) items.push('<span class="gap">…</span>');
      items.push('<button type="button" data-dep-page="' + p + '" class="' + (p === cur ? 'is-active' : '') + '">' + p + '</button>');
      prev = p;
    });
    wrap.innerHTML =
      '<button type="button" data-dep-page="' + (cur - 1) + '" ' + (cur === 1 ? 'disabled' : '') + ' aria-label="Sebelumnya"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="m14 6-6 6 6 6"/></svg></button>' +
      items.join('') +
      '<button type="button" data-dep-page="' + (cur + 1) + '" ' + (cur === pages ? 'disabled' : '') + ' aria-label="Berikutnya"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="m10 6 6 6-6 6"/></svg></button>';
  }

  function openDepDetail(ref) {
    var a = ASET.find(function (x) { return x.ref === ref; });
    if (!a) return;
    var d = depOf(a);
    var l = curLoc(a);
    var html = '<div class="modal-head">' +
      '<div><h3>Detail Penyusutan</h3><p>Metode garis lurus · residu Rp 0</p></div>' +
      '<button class="icon-btn modal-close" data-close aria-label="Tutup"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="m6 6 12 12M18 6 6 18"/></svg></button>' +
      '</div><div class="modal-body">' +
      '<div class="detail-hero"><div class="dh-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M4 5v14h16"/><path d="m7 15 4-4 3 3 4-6"/><path d="M15 8h3v3"/></svg></div>' +
      '<div><h4>' + esc(a.name) + '</h4><p>Kode aset ' + esc(a.ref) + ' · ' + esc(locFull(l)) + '</p></div></div>' +
      '<div class="detail-grid">' +
        detailItem('Kategori Aset', katNameOf(a), false, true) +
        detailItem('Tanggal Perolehan', fmtDate(a.date)) +
        detailItem('Harga Perolehan', rp(a.price)) +
        detailItem('Nilai Residu', 'Rp 0') +
        detailItem('Masa Manfaat', d.years ? d.years + ' tahun (' + d.lifeMonths + ' bulan)' : '—') +
        detailItem('Umur Aset', d.years ? formatAge(d.ageMonths) : '—') +
        detailItem('Akumulasi Penyusutan', d.years ? rp(d.accumulated) : '—') +
        detailItem('Nilai Buku', d.years ? rp(d.book) : '—') +
      '</div>' +
      '<div class="receipt">' +
        receiptRow('Penyusutan tahunan', d.years ? rp(Math.round(d.annual)) : '—') +
        receiptRow('Penyusutan bulanan', d.years ? rp(Math.round(d.monthly)) : '—') +
        receiptRow('Rumus', d.years ? '(' + rp(a.price) + ' − 0) ÷ ' + d.years + ' tahun' : 'Kategori belum dikenali') +
      '</div>' +
      '<p class="tips-note">Penyusutan dihitung sejak tanggal perolehan dan berhenti setelah masa manfaat habis; nilai buku tidak menjadi negatif.</p>' +
      '</div><div class="modal-foot"><button class="btn btn-ghost" data-close>Tutup</button>' +
      '<button class="btn btn-primary" data-goto="aset">Lihat Data Aset</button></div>';
    openModal(html);
  }

  function exportDepreciation() {
    var rows = [['Kode Aset', 'Nama Barang', 'Kategori', 'Masa Manfaat (tahun)', 'Tanggal Perolehan', 'Harga Satuan',
      'Penyusutan Tahunan', 'Penyusutan Bulanan', 'Umur (bulan)', 'Akumulasi Penyusutan', 'Nilai Buku', 'Status']];
    filteredDep().forEach(function (x) {
      var d = x.d;
      rows.push([
        x.a.ref, x.a.name, katNameOf(x.a), d.years || '', x.a.date, x.a.price,
        Math.round(d.annual), Math.round(d.monthly), d.ageMonths, d.accumulated, d.book,
        !d.years ? 'Tanpa kategori' : (d.done ? 'Selesai' : 'Berjalan')
      ]);
    });
    downloadCSV('penyusutan-garis-lurus-' + todayISO() + '.csv', rows);
    toast(rows.length - 1 + ' aset diekspor ke CSV.');
  }

  /* ---------------- Ekspor ---------------- */
  function downloadCSV(filename, rows) {
    var csv = rows.map(function (r) {
      return r.map(function (c) {
        var s = String(c == null ? '' : c);
        return /[";\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
      }).join(';');
    }).join('\r\n');
    var blob = new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8;' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
  }

  function exportAssets() {
    var rows = [['Kode Aset', 'Nama Barang', 'Sub Kelompok', 'Lokasi', 'Kode Lokasi', 'Tanggal Perolehan', 'Merek / Detail', 'Nilai Perolehan']];
    filteredAssets().forEach(function (a) {
      var l = curLoc(a);
      rows.push([a.ref, a.name, titleCase(subName(a.sub)), locRaw(l)[1], locCode(l), a.date, brandName(a) || a.desc, a.price]);
    });
    downloadCSV('data-aset-' + todayISO() + '.csv', rows);
    toast(rows.length - 1 + ' aset diekspor ke CSV.');
  }

  function exportHistory() {
    var rows = [['Tanggal', 'Kode Aset', 'Nama Aset', 'Dari', 'Ke', 'Petugas', 'Alasan', 'Catatan']];
    historyRows().forEach(function (t) {
      rows.push([t.date, t.ref, t.name, locRaw(t.from)[1], locRaw(t.to)[1], t.pic, t.reason, t.note]);
    });
    downloadCSV('riwayat-perpindahan-' + todayISO() + '.csv', rows);
    toast(rows.length - 1 + ' perpindahan diekspor ke CSV.');
  }

  /* ---------------- Admin: PIN & log ---------------- */
  function pinHash(pin) {
    var h = 0x811c9dc5;
    var s = 'simpati:' + String(pin);
    for (var i = 0; i < s.length; i++) {
      h ^= s.charCodeAt(i);
      h = (h * 16777619) >>> 0;
    }
    return h.toString(36);
  }

  function storedPinHash() {
    try { return localStorage.getItem(ADMIN_PIN_KEY) || pinHash('1234'); } catch (e) { return pinHash('1234'); }
  }

  function isAdminSession() {
    try { return sessionStorage.getItem(ADMIN_SESSION_KEY) === '1'; } catch (e) { return false; }
  }

  function loadAdminLog() {
    try {
      var raw = localStorage.getItem(ADMIN_LOG_KEY);
      var parsed = raw ? JSON.parse(raw) : null;
      admin.log = Array.isArray(parsed) ? parsed.slice(0, 40) : [];
    } catch (e) { admin.log = []; }
  }

  function saveAdminLog() {
    try { localStorage.setItem(ADMIN_LOG_KEY, JSON.stringify(admin.log.slice(0, 40))); } catch (e) { /* abaikan */ }
  }

  function addLog(action, detail) {
    admin.log.unshift({ at: new Date().toISOString(), action: action, detail: detail });
    admin.log = admin.log.slice(0, 40);
    saveAdminLog();
    if (admin.unlocked && ui.view === 'admin') renderAdminLog();
  }

  function renderAdminLog() {
    var box = $('#admin-log');
    if (!box) return;
    if (!admin.log.length) {
      box.innerHTML = '<div class="log-empty">Belum ada aktivitas admin pada perangkat ini.</div>';
      return;
    }
    box.innerHTML = admin.log.map(function (l) {
      return '<div class="log-item"><span class="log-dot"></span><div class="log-body"><strong>' + esc(l.action) + '</strong><p>' + esc(l.detail || '') + '</p></div><time>' + esc(fmtDateTime(l.at)) + '</time></div>';
    }).join('');
  }

  function lockAdmin() {
    admin.unlocked = false;
    try { sessionStorage.removeItem(ADMIN_SESSION_KEY); } catch (e) { /* abaikan */ }
    addLog('Kunci panel', 'Sesi admin ditutup');
    renderAdmin();
    toast('Panel admin dikunci.');
  }

  function renderAdmin() {
    var gate = $('#admin-gate');
    var panel = $('#admin-panel');
    if (!gate || !panel) return;

    $('#admin-head-actions').innerHTML = admin.unlocked
      ? '<button class="btn btn-ghost" data-lock-admin><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7"><rect x="5" y="10" width="14" height="10" rx="2"/><path d="M8 10V8a4 4 0 0 1 8 0v2"/></svg> Kunci panel</button>' +
        '<button class="btn btn-primary" data-goto="tambah"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M12 5v14M5 12h14"/></svg> Tambah Barang</button>'
      : '<span class="badge badge-neutral"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7"><rect x="5" y="10" width="14" height="10" rx="2"/><path d="M8 10V8a4 4 0 0 1 8 0v2"/></svg> Panel terkunci</span>';

    gate.hidden = admin.unlocked;
    panel.hidden = !admin.unlocked;

    if (!admin.unlocked) {
      if ($('#admin-pin')) $('#admin-pin').value = '';
      if ($('#admin-login-error')) $('#admin-login-error').hidden = true;
      return;
    }
    renderAdminKpis();
    renderAdminDataset();
    renderAdminLog();
    if ($('#admin-pin-form')) $('#admin-pin-form').reset();
    renderUploadUI();
  }

  function adminStat(label, value, unit, icon, color) {
    var paths = {
      layers: '<path d="m12 3 8 4.5-8 4.5-8-4.5L12 3Z"/><path d="m4 12 8 4.5 8-4.5"/><path d="m4 16.5 8 4.5 8-4.5"/>',
      wallet: '<path d="M4 7.5A2.5 2.5 0 0 1 6.5 5H19a1 1 0 0 1 1 1v12a1 1 0 0 1-1 1H6.5A2.5 2.5 0 0 1 4 16.5v-9Z"/><path d="M4 8h13"/><path d="M16 12.5h2"/>',
      pin: '<path d="M12 21s-6-5.4-6-10a6 6 0 0 1 12 0c0 4.6-6 10-6 10Z"/><circle cx="12" cy="11" r="2.5"/>',
      file: '<path d="M6 3h8l4 4v14a1 1 0 0 1-1 1H6a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1Z"/><path d="M14 3v4h4M8 12h8M8 16h5"/>'
    };
    return '<article class="admin-stat" style="--stat-color:' + color + '">' +
      '<span class="admin-stat-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round">' + paths[icon] + '</svg></span>' +
      '<div><span>' + esc(label) + '</span><strong>' + esc(value) + (unit ? ' <small>' + esc(unit) + '</small>' : '') + '</strong></div></article>';
  }

  function renderAdminKpis() {
    var total = ASET.reduce(function (s, a) { return s + a.price; }, 0);
    var kindLabel = DATASET.kind === 'default' ? 'Bawaan' : (DATASET.kind === 'upload' ? 'Excel' : 'Manual');
    $('#admin-kpis').innerHTML =
      adminStat('Jumlah Aset', nf.format(ASET.length), 'unit', 'layers', '#0d9488') +
      adminStat('Total Nilai', rupiahShortText(total), '', 'wallet', '#2f7de1') +
      adminStat('Lokasi Aktif', nf.format(uniqueLocations()), 'ruang', 'pin', '#8b5cf6') +
      adminStat('Sumber Data', kindLabel, '', 'file', '#e59a17');
  }

  function renderAdminDataset() {
    var box = $('#admin-dataset');
    if (!box) return;
    var kindLabel = DATASET.kind === 'default' ? 'Data bawaan' : (DATASET.kind === 'upload' ? 'Unggahan Excel' : 'Data manual');
    var kindBadge = DATASET.kind === 'default' ? 'badge-neutral' : (DATASET.kind === 'upload' ? 'badge-teal' : 'badge-violet');
    var updated = DATASET.updatedAt ? fmtDateTime(DATASET.updatedAt) : '—';
    var fileLine = DATASET.fileName ? DATASET.fileName + (DATASET.sheet ? ' · sheet ' + DATASET.sheet : '') : '—';

    box.innerHTML =
      '<div class="card-head"><div><h3>Dataset Aktif</h3><p>Sumber data yang sedang dipakai aplikasi</p></div><span class="badge ' + kindBadge + '">' + esc(kindLabel) + '</span></div>' +
      '<div class="dataset-body">' +
        '<div class="dataset-row"><span>Sumber</span><strong title="' + esc(DATASET.source) + '">' + esc(DATASET.source || '—') + '</strong></div>' +
        '<div class="dataset-row"><span>File</span><strong title="' + esc(fileLine) + '">' + esc(fileLine) + '</strong></div>' +
        '<div class="dataset-row"><span>Jumlah aset</span><strong>' + nf.format(ASET.length) + ' unit</strong></div>' +
        '<div class="dataset-row"><span>Lokasi</span><strong>' + nf.format(uniqueLocations()) + ' ruang</strong></div>' +
        '<div class="dataset-row"><span>Total nilai</span><strong>' + rupiahShortText(ASET.reduce(function (s, a) { return s + a.price; }, 0)) + '</strong></div>' +
        '<div class="dataset-row"><span>Terakhir diubah</span><strong>' + esc(updated) + '</strong></div>' +
        '<div class="dataset-row"><span>Penyimpanan</span><strong class="' + (window.DatasetStore ? 'ok' : 'warn') + '">' + (window.DatasetStore && window.DatasetStore.supported() ? 'IndexedDB aktif' : 'Memori sesi') + '</strong></div>' +
      '</div>' +
      '<div class="dataset-foot"><button type="button" class="link-btn" data-goto="aset">Lihat Data Aset</button></div>';
  }

  function loginAdmin(pin) {
    if (pinHash(pin) !== storedPinHash()) return false;
    admin.unlocked = true;
    try { sessionStorage.setItem(ADMIN_SESSION_KEY, '1'); } catch (e) { /* abaikan */ }
    addLog('Masuk panel admin', 'Akses admin dibuka');
    renderAdmin();
    toast('Selamat datang di panel admin.');
    return true;
  }

  /* ---------------- Admin: unggah Excel ---------------- */
  var UPLOAD_FIELDS = [
    { key: 'ref', label: 'Kode Aset', required: true, keys: ['ref', 'kode aset', 'kode barang', 'kode aktiva', 'kode', 'no aset', 'asset id'], exclude: /dokumen/ },
    { key: 'name', label: 'Nama Barang', required: true, keys: ['nama barang', 'barang', 'nama aset', 'nama', 'item'], exclude: /kelompok|sub|tipe|kategori/ },
    { key: 'lokasi', label: 'Lokasi / Ruang', required: true, keys: ['lokasi', 'ruang', 'unit kerja', 'tempat', 'posisi', 'location'] },
    { key: 'date', label: 'Tanggal Perolehan', keys: ['tanggal perolehan', 'tanggal', 'tgl perolehan', 'tgl', 'perolehan', 'acquisition date'], exclude: /dokumen/ },
    { key: 'sub', label: 'Sub Kelompok', required: true, keys: ['sub kelompok barang', 'sub kelompok', 'subkelompok', 'sub kategori', 'subkategori'] },
    { key: 'kel', label: 'Kelompok', keys: ['kelompok barang', 'kelompok', 'golongan', 'kategori'], exclude: /sub/ },
    { key: 'brand', label: 'Merek / Detail', keys: ['merek', 'merk', 'brand', 'detail', 'model'] },
    { key: 'desc', label: 'Deskripsi', keys: ['deskripsi', 'keterangan', 'catatan', 'uraian'] },
    { key: 'price', label: 'Harga / Nilai', required: true, keys: ['harga satuan', 'harga', 'nilai perolehan', 'nilai', 'price', 'amount', 'biaya'] }
  ];

  var UPLOAD_ORDER = ['sub', 'kel', 'date', 'price', 'lokasi', 'ref', 'desc', 'brand', 'name'];

  function normHeader(value) {
    return cleanText(value).toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
  }

  function autoMap(headers, sampleRows) {
    var used = {};
    var mapData = {};
    UPLOAD_ORDER.forEach(function (key) {
      var field = UPLOAD_FIELDS.find(function (f) { return f.key === key; });
      var best = -1;
      var bestScore = 0;
      headers.forEach(function (h, i) {
        if (used[i]) return;
        var hh = normHeader(h);
        if (!hh) return;
        if (field.exclude && field.exclude.test(hh)) return;
        var score = 0;
        field.keys.forEach(function (k) {
          if (hh === k) score = Math.max(score, 6);
          else if (hh.indexOf(k) === 0) score = Math.max(score, 4);
          else if (hh.indexOf(k) !== -1) score = Math.max(score, 2);
        });
        if (score > bestScore) { bestScore = score; best = i; }
      });
      if (best >= 0) { mapData[key] = best; used[best] = 1; }
    });

    /* Kolom lokasi kadang tidak berjudul. Kenali dari isi selnya. */
    if (mapData.lokasi === undefined && sampleRows && sampleRows.length) {
      var sample = sampleRows.slice(0, 80);
      var bestCol = -1;
      var bestRatio = 0;
      for (var c = 0; c < headers.length; c++) {
        if (used[c]) continue;
        var filled = 0;
        var hits = 0;
        for (var r = 0; r < sample.length; r++) {
          var v = cleanText((sample[r] || [])[c]);
          if (!v) continue;
          filled++;
          if (/^(?:lokasi|lok\.?)\s*[0-9]/i.test(v) || /^[0-9]{1,4}[.\-]\s*[0-9]+\s*[-–—]/.test(v)) hits++;
        }
        var ratio = filled >= 3 ? hits / filled : 0;
        if (ratio >= 0.6 && ratio > bestRatio) { bestRatio = ratio; bestCol = c; }
      }
      if (bestCol >= 0) mapData.lokasi = bestCol;
    }
    return mapData;
  }

  function detectHeaderRow(rows) {
    var limit = Math.min(20, rows.length);
    for (var i = 0; i < limit; i++) {
      var row = rows[i] || [];
      var nonEmpty = row.filter(function (v) { return cleanText(v) !== ''; }).length;
      var text = row.map(function (v) { return cleanText(v); }).join(' ').toLowerCase();
      if (nonEmpty >= 3 && /barang|nama|kode|uraian|asset|aset|ref/.test(text)) return i + 1;
    }
    for (var j = 0; j < rows.length; j++) {
      if ((rows[j] || []).some(function (v) { return cleanText(v) !== ''; })) return j + 1;
    }
    return 1;
  }

  function colLetter(index) {
    var n = index + 1;
    var s = '';
    while (n > 0) {
      var m = (n - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  function parseLocation(v) {
    var s = cleanText(v);
    if (!s) return { code: '', name: 'TANPA LOKASI' };
    var m = s.match(/^(?:lokasi|lok\.?)\s*([0-9][0-9.]*)\s*[-–—:]\s*(.+)$/i);
    if (m) return { code: cleanText(m[1]), name: cleanText(m[2]) };
    var m2 = s.match(/^([0-9][0-9.]*)\s*[-–—]\s*(.+)$/);
    if (m2) return { code: cleanText(m2[1]), name: cleanText(m2[2]) };
    return { code: '', name: s };
  }

  function buildDbFromRows(rows, mapData) {
    var lokasi = [];
    var locKey = {};
    var sub = [];
    var subKey = {};
    var kel = [];
    var kelKey = {};
    var merek = [];
    var merkKey = {};
    var aset = [];
    var stats = { rows: 0, valid: 0, skipped: 0, duplicates: 0, issues: [], totalValue: 0, locations: 0, subs: 0 };
    var refSeen = {};

    function addLookup(arr, map, value, key) {
      if (map[key] !== undefined) return map[key];
      var id = arr.length;
      arr.push(value);
      map[key] = id;
      return id;
    }

    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      if (!row || !row.length) continue;
      var any = false;
      for (var c = 0; c < row.length; c++) {
        if (cleanText(row[c])) { any = true; break; }
      }
      if (!any) continue;
      stats.rows++;

      function cell(field) {
        var idx = mapData[field];
        if (idx === '' || idx === undefined || idx === null) return '';
        return row[+idx];
      }

      var ref = cleanText(cell('ref'));
      var name = cleanText(cell('name'));
      if (/^(total|jumlah|grand total|sub total|subtotal)\b/i.test(ref) || /^(total|jumlah|grand total|sub total|subtotal)\b/i.test(name)) {
        stats.skipped++;
        continue;
      }
      if (!ref || !name) {
        stats.skipped++;
        if (stats.issues.length < 6) {
          stats.issues.push('Baris data ' + (i + 1) + ': ' + (!ref ? 'kode aset kosong' : 'nama barang kosong') + ' — dilewati.');
        }
        continue;
      }

      var loc = parseLocation(cell('lokasi'));
      var locKeyValue = loc.code ? ('C:' + loc.code.toUpperCase()) : ('N:' + loc.name.toUpperCase());
      var locId = addLookup(lokasi, locKey, [loc.code, loc.name], locKeyValue);

      var subValue = cleanText(cell('sub')).toUpperCase() || 'LAIN-LAIN';
      var subId = addLookup(sub, subKey, subValue, subValue);

      var kelValue = cleanText(cell('kel')).toUpperCase() || 'AKTIVA TETAP BERWUJUD';
      var kelId = addLookup(kel, kelKey, kelValue, kelValue);

      var brandValue = cleanText(cell('brand'));
      var brandId = -1;
      if (brandValue) {
        var brandKeyValue = brandValue.toLowerCase();
        brandId = addLookup(merek, merkKey, brandValue, brandKeyValue);
      }

      var price = parseMoneyValue(cell('price'));
      var date = parseDateValue(cell('date'));
      var desc = cleanText(cell('desc'));

      var refKey = ref.toUpperCase();
      if (refSeen[refKey]) stats.duplicates++;
      refSeen[refKey] = 1;

      aset.push([ref, name, locId, date, subId, brandId, price, kelId, desc]);
      stats.valid++;
      stats.totalValue += price;
    }

    if (!lokasi.length) {
      lokasi.push(['', 'TANPA LOKASI']);
      locKey['N:TANPA LOKASI'] = 0;
    }
    if (!sub.length) { sub.push('LAIN-LAIN'); subKey['LAIN-LAIN'] = 0; }
    if (!kel.length) { kel.push('AKTIVA TETAP BERWUJUD'); kelKey['AKTIVA TETAP BERWUJUD'] = 0; }

    stats.locations = lokasi.length;
    stats.subs = sub.length;

    var db = {
      lokasi: lokasi, sub: sub, kel: kel, merek: merek, aset: aset,
      meta: { total: stats.totalValue, count: aset.length }
    };
    return { db: db, stats: stats };
  }

  function mergeLookup(baseArr, addArr, keyFn) {
    var map = {};
    baseArr.forEach(function (v, i) {
      var k = keyFn(v);
      if (map[k] === undefined) map[k] = i;
    });
    var remap = {};
    addArr.forEach(function (v, i) {
      var k = keyFn(v);
      if (map[k] === undefined) {
        map[k] = baseArr.length;
        baseArr.push(v);
      }
      remap[i] = map[k];
    });
    return remap;
  }

  function mergeDb(base, add) {
    var db = cloneDb(base);
    var locMap = mergeLookup(db.lokasi, add.lokasi, function (l) { return String(l[0] || '').toUpperCase() + '|' + String(l[1] || '').toUpperCase(); });
    var subMap = mergeLookup(db.sub, add.sub, function (s) { return String(s).toUpperCase(); });
    var kelMap = mergeLookup(db.kel, add.kel, function (s) { return String(s).toUpperCase(); });
    var brandMap = mergeLookup(db.merek, add.merek, function (s) { return String(s).toLowerCase(); });

    var incoming = {};
    add.aset.forEach(function (a) { incoming[String(a[0]).toUpperCase()] = 1; });
    db.aset = db.aset.filter(function (a) { return !incoming[String(a[0]).toUpperCase()]; });
    add.aset.forEach(function (a) {
      db.aset.push([
        a[0], a[1],
        a[2] >= 0 ? locMap[a[2]] : a[2],
        a[3],
        a[4] >= 0 ? subMap[a[4]] : a[4],
        a[5] >= 0 ? brandMap[a[5]] : a[5],
        a[6],
        a[7] >= 0 ? kelMap[a[7]] : a[7],
        a[8]
      ]);
    });
    return db;
  }

  function uploadReset() {
    upload = {
      step: 1, fileName: '', fileSize: 0, fileType: '', sheets: [], sheetIndex: 0,
      sheetName: '', rows: [], headerRow: 1, headers: [], map: {}, built: null,
      mode: 'replace', pending: null, success: null, error: '', loading: false
    };
    if ($('#upload-file')) $('#upload-file').value = '';
    renderUploadUI();
  }

  function uploadError(msg) {
    upload.error = msg || '';
    upload.loading = false;
    upload.step = 1;
    renderUploadUI();
  }

  function fmtBytes(n) {
    if (n >= 1048576) return (n / 1048576).toFixed(1).replace('.', ',') + ' MB';
    if (n >= 1024) return Math.round(n / 1024) + ' KB';
    return n + ' B';
  }

  function readUploadFile(file) {
    return new Promise(function (resolve, reject) {
      var ext = file.name.split('.').pop().toLowerCase();
      var reader = new FileReader();
      reader.onerror = function () { reject(new Error('File tidak dapat dibaca.')); };
      if (ext === 'csv') {
        reader.onload = function () {
          var rows = window.XLSXReader.parseDelimited(String(reader.result || ''));
          resolve({ sheets: [{ name: 'CSV', rows: rows }] });
        };
        reader.readAsText(file);
      } else {
        reader.onload = function () {
          try {
            var wb = window.XLSXReader.read(reader.result);
            resolve(wb);
          } catch (err) {
            reject(err);
          }
        };
        reader.readAsArrayBuffer(file);
      }
    });
  }

  function pickSheetIndex(sheets) {
    var best = 0;
    var bestScore = -1;
    sheets.forEach(function (s, i) {
      var score = s.rows.length;
      if (/all|data|aset|daftar/i.test(s.name)) score += 100000;
      if (score > bestScore) { bestScore = score; best = i; }
    });
    return best;
  }

  function handleUploadFile(file) {
    if (!file) return;
    upload.error = '';
    upload.success = null;
    var ext = file.name.split('.').pop().toLowerCase();
    if (ext !== 'xlsx' && ext !== 'csv') {
      uploadError('Format file tidak didukung. Gunakan .xlsx (Excel modern) atau .csv.');
      return;
    }
    if (file.size > 20 * 1024 * 1024) {
      uploadError('File terlalu besar. Batas maksimum 20 MB.');
      return;
    }

    upload.fileName = file.name;
    upload.fileSize = file.size;
    upload.fileType = ext;
    upload.loading = true;
    upload.rows = [];
    upload.sheets = [];
    upload.built = null;
    upload.step = 1;
    renderUploadUI();

    readUploadFile(file).then(function (wb) {
      if (!wb || !wb.sheets || !wb.sheets.length) throw new Error('Tidak ada data pada file.');
      upload.sheets = wb.sheets;
      upload.sheetIndex = pickSheetIndex(wb.sheets);
      upload.sheetName = wb.sheets[upload.sheetIndex].name;
      upload.rows = wb.sheets[upload.sheetIndex].rows;
      upload.headerRow = detectHeaderRow(upload.rows);
      upload.loading = false;
      applyUploadHeaders();
      upload.step = 2;
      renderUploadUI();
    }).catch(function (err) {
      uploadError((err && err.message) || 'File tidak dapat diproses.');
    });
  }

  function applyUploadHeaders() {
    if (!upload.rows.length) return;
    var limit = Math.min(upload.rows.length, upload.headerRow > 0 ? upload.headerRow : 1);
    var row = upload.rows[limit - 1] || [];
    upload.headers = [];
    for (var i = 0; i < row.length; i++) {
      upload.headers[i] = cleanText(row[i]) || 'Kolom ' + colLetter(i);
    }
    upload.map = autoMap(upload.headers, upload.rows.slice(limit, limit + 80));
    refreshUpload();
  }

  function refreshUpload() {
    if (!upload.rows.length) return;
    var start = Math.min(upload.rows.length, upload.headerRow);
    upload.built = buildDbFromRows(upload.rows.slice(start), upload.map);
    renderUploadPreview();
    renderUploadStats();
  }

  function renderUploadUI() {
    if (!$('#upload-step-file')) return;
    var step = upload.step || 1;

    $('#upload-step-file').hidden = step !== 1;
    $('#upload-step-map').hidden = step !== 2;
    $('#upload-step-review').hidden = step !== 3;
    $('#upload-step-badge').textContent = '0' + step + ' / 03';
    $$('.ustep').forEach(function (el) {
      var n = +el.getAttribute('data-ustep');
      el.classList.toggle('is-active', n === step);
      el.classList.toggle('is-done', n < step);
    });

    var err = $('#upload-error');
    err.hidden = !upload.error;
    err.textContent = upload.error || '';

    var chip = $('#upload-file-chip');
    if (upload.fileName) {
      chip.hidden = false;
      chip.innerHTML =
        '<span class="file-chip-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><path d="M6 3h8l4 4v14a1 1 0 0 1-1 1H6a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1Z"/><path d="M14 3v4h4M8 12h8M8 16h5"/></svg></span>' +
        '<span class="file-chip-body"><strong>' + esc(upload.fileName) + '</strong><small>' + (upload.loading ? 'Membaca file…' : fmtBytes(upload.fileSize) + ' · ' + upload.sheets.length + ' sheet') + '</small></span>' +
        '<button type="button" class="link-btn" data-upload-remove>Ganti file</button>';
    } else {
      chip.hidden = true;
      chip.innerHTML = '';
    }

    var success = $('#upload-success');
    if (upload.success) {
      success.hidden = false;
      success.innerHTML =
        '<div class="success-hero"><div class="success-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M5 12.5 10 17l9-10"/></svg></div>' +
        '<h3>Data berhasil diimpor</h3><p>' + esc(upload.success.summary) + '</p><small>' + esc(upload.success.file) + '</small></div>' +
        '<div class="upload-success-actions"><button class="btn btn-ghost" data-goto="aset">Lihat Data Aset</button><button class="btn btn-primary" data-upload-again>Unggah file lain</button></div>';
    } else {
      success.hidden = true;
      success.innerHTML = '';
    }

    if (step !== 2 && step !== 3) return;

    var sheetSel = $('#upload-sheet');
    sheetSel.innerHTML = upload.sheets.map(function (s, i) {
      return '<option value="' + i + '">' + esc(s.name) + ' (' + nf.format(s.rows.length) + ' baris)</option>';
    }).join('');
    sheetSel.value = String(upload.sheetIndex);
    $('#upload-header').max = upload.rows.length;
    $('#upload-header').value = upload.headerRow;
    $('#upload-file-info').textContent = upload.fileName + ' · ' + upload.sheets.length + ' sheet · ' +
      nf.format(upload.rows.length) + ' baris terbaca · header di baris ' + upload.headerRow;

    renderUploadMap();
    renderUploadPreview();
    renderUploadStats();
    renderUploadReview();
  }

  function renderUploadMap() {
    var box = $('#upload-map');
    if (!box) return;
    box.innerHTML = UPLOAD_FIELDS.map(function (f) {
      var options = '<option value="">— tidak dipakai —</option>' + upload.headers.map(function (h, i) {
        return '<option value="' + i + '"' + (upload.map[f.key] === i ? ' selected' : '') + '>' + colLetter(i) + ' · ' + esc(h) + '</option>';
      }).join('');
      return '<div class="map-item"><label for="map-' + f.key + '">' + esc(f.label) + (f.required ? ' <span class="req">*</span>' : '') + '</label>' +
        '<select id="map-' + f.key + '" class="select select-full" data-map="' + f.key + '">' + options + '</select></div>';
    }).join('');
  }

  function renderUploadPreview() {
    var box = $('#upload-preview');
    if (!box) return;
    if (!upload.built || !upload.built.db.aset.length) {
      box.innerHTML = '<p class="preview-empty-line">Belum ada baris valid untuk ditampilkan.</p>';
      return;
    }
    var db = upload.built.db;
    var rows = db.aset.slice(0, 5);
    box.innerHTML = '<table class="preview-table"><thead><tr><th>Kode</th><th>Nama Barang</th><th>Sub Kelompok</th><th>Lokasi</th><th>Tanggal</th><th class="num">Harga</th></tr></thead><tbody>' +
      rows.map(function (a) {
        return '<tr><td><span class="cell-kode">' + esc(a[0]) + '</span></td>' +
          '<td>' + esc(a[1]) + '</td>' +
          '<td>' + esc(a[4] >= 0 && db.sub[a[4]] ? titleCase(db.sub[a[4]]) : '—') + '</td>' +
          '<td>' + esc(a[2] >= 0 && db.lokasi[a[2]] ? titleCase(db.lokasi[a[2]][1].replace(/^RSH MADIUN\s+/i, '')) : '—') + '</td>' +
          '<td>' + (a[3] ? fmtDate(a[3]) : '—') + '</td>' +
          '<td class="num">' + rp(a[6]) + '</td></tr>';
      }).join('') + '</tbody></table>';
  }

  function renderUploadStats() {
    var statsBox = $('#upload-stats');
    var issuesBox = $('#upload-issues');
    if (!statsBox) return;
    if (!upload.built) {
      statsBox.innerHTML = '';
      if (issuesBox) issuesBox.innerHTML = '';
      return;
    }
    var s = upload.built.stats;
    statsBox.innerHTML =
      '<div class="import-stat"><strong>' + nf.format(s.valid) + '</strong><span>baris valid</span></div>' +
      '<div class="import-stat"><strong>' + nf.format(s.skipped) + '</strong><span>baris dilewati</span></div>' +
      '<div class="import-stat"><strong>' + nf.format(s.duplicates) + '</strong><span>kode duplikat</span></div>' +
      '<div class="import-stat"><strong>' + nf.format(s.locations) + '</strong><span>lokasi</span></div>' +
      '<div class="import-stat wide"><strong>' + rupiahShortText(s.totalValue) + '</strong><span>total nilai aset</span></div>';

    var missing = UPLOAD_FIELDS.filter(function (f) {
      return f.required && (upload.map[f.key] === '' || upload.map[f.key] === undefined);
    }).map(function (f) { return f.label; });

    var notes = [];
    if (missing.length) notes.push('<div class="issue issue-error"><strong>Kolom wajib belum dipetakan</strong><p>' + esc(missing.join(', ')) + '. Pilih kolom yang sesuai sebelum mengimpor.</p></div>');
    if (s.valid === 0) notes.push('<div class="issue issue-error"><strong>Tidak ada baris valid</strong><p>Periksa pemetaan kolom dan baris header.</p></div>');
    if (s.skipped > 0 && s.issues.length) notes.push('<div class="issue issue-warn"><strong>' + nf.format(s.skipped) + ' baris dilewati</strong><p>' + esc(s.issues.join(' ')) + '</p></div>');
    if (s.duplicates > 0) notes.push('<div class="issue issue-warn"><strong>' + nf.format(s.duplicates) + ' kode duplikat</strong><p>Kode yang sama tetap diimpor; gunakan mode gabungkan untuk memperbarui data per kode.</p></div>');
    if (!missing.length && s.valid > 0) notes.push('<div class="issue issue-ok"><strong>Data siap diimpor</strong><p>' + nf.format(s.valid) + ' aset akan ' + (upload.mode === 'merge' ? 'digabungkan' : 'menggantikan data saat ini') + '.</p></div>');

    if (issuesBox) issuesBox.innerHTML = notes.join('');

    var btn = $('#btn-do-import');
    if (btn) {
      var disabled = missing.length > 0 || s.valid === 0;
      btn.disabled = disabled;
      btn.classList.toggle('is-disabled', disabled);
    }
  }

  function renderUploadReview() {
    var mode = $('#upload-mode');
    if (!mode) return;
    $$('input[name="import-mode"]').forEach(function (r) { r.checked = r.value === upload.mode; });
  }

  function confirmImport() {
    var built = upload.built;
    if (!built || !built.db.aset.length) return;
    var missing = UPLOAD_FIELDS.filter(function (f) {
      return f.required && (upload.map[f.key] === '' || upload.map[f.key] === undefined);
    });
    if (missing.length) { toast('Lengkapi pemetaan kolom wajib terlebih dahulu.'); return; }

    var mode = upload.mode;
    var baseDb = mode === 'merge' ? mergeDb(DB, built.db) : built.db;
    var info = {
      kind: 'upload',
      source: upload.fileName ? upload.fileName.replace(/\.(xlsx|csv)$/i, '') : 'Data Excel',
      fileName: upload.fileName,
      sheet: upload.sheetName,
      updatedAt: new Date().toISOString(),
      importedAt: new Date().toISOString()
    };
    withMeta(baseDb, info);

    var refs = {};
    baseDb.aset.forEach(function (a) { refs[String(a[0])] = 1; });
    var lost = state.transfers.filter(function (t) { return !refs[t.ref]; }).length;

    upload.pending = { db: baseDb, info: info, stats: built.stats, mode: mode };
    openModal(
      '<div class="modal-head"><div><h3>Impor data aset?</h3><p>Periksa ringkasan sebelum mengganti data.</p></div>' +
      '<button class="icon-btn modal-close" data-close aria-label="Tutup"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="m6 6 12 12M18 6 6 18"/></svg></button></div>' +
      '<div class="modal-body"><div class="receipt">' +
        receiptRow('File', upload.fileName) +
        receiptRow('Sheet', upload.sheetName) +
        receiptRow('Aset valid', nf.format(built.stats.valid) + ' unit') +
        receiptRow('Baris dilewati', nf.format(built.stats.skipped)) +
        receiptRow('Total nilai', rupiahShortText(built.stats.totalValue)) +
        receiptRow('Mode', mode === 'merge' ? 'Gabungkan (upsert)' : 'Ganti seluruh data') +
      '</div>' +
      (mode === 'replace' ? '<p class="modal-warn">Data aset saat ini akan diganti. Lokasi aset dihitung ulang dari riwayat yang tersisa.</p>' : '') +
      '<p class="modal-warn">' + nf.format(lost) + ' catatan perpindahan yang asetnya tidak ada pada data baru akan dihapus.</p>' +
      '</div><div class="modal-foot"><button class="btn btn-ghost" data-close>Batal</button>' +
      '<button class="btn btn-primary" data-confirm-import>Ya, impor data</button></div>'
    );
  }

  function finishImport() {
    var p = upload.pending;
    if (!p) return;
    applyDataset(p.db, p.info);
    persistDataset();
    addLog('Impor data Excel', p.info.fileName + ' · ' + nf.format(p.stats.valid) + ' aset · mode ' + (p.mode === 'merge' ? 'gabung' : 'ganti'));
    upload.success = {
      summary: nf.format(p.stats.valid) + ' aset · ' + nf.format(p.stats.locations) + ' lokasi · ' + rupiahShortText(p.stats.totalValue),
      file: p.info.fileName + ' · sheet ' + p.info.sheet + ' · mode ' + (p.mode === 'merge' ? 'Gabungkan' : 'Ganti') + ' · disimpan permanen'
    };
    upload.step = 1;
    upload.fileName = '';
    upload.fileSize = 0;
    upload.sheets = [];
    upload.rows = [];
    upload.built = null;
    upload.pending = null;
    upload.error = '';
    upload.loading = false;
    if ($('#upload-file')) $('#upload-file').value = '';
    closeModal();
    renderAdmin();
    toast('Data aset berhasil diimpor.');
  }

  function downloadTemplate() {
    var rows = [
      ['No', 'Ref', 'Barang', 'Lokasi', 'Tanggal Perolehan', 'Kelompok Barang', 'Sub Kelompok Barang', 'Deskripsi', 'Detail', 'Harga Satuan'],
      ['1', '1764419', 'KURSI TUNGGU', 'Lokasi 049.01 - AREA PARKIR LANTAI 1', '2024-11-01', 'AKTIVA TETAP BERWUJUD', 'ALAT UMUM DAN RUMAH TANGGA', 'Kursi tunggu ruang tunggu', '', '1250000']
    ];
    downloadCSV('template-data-aset.csv', rows);
    toast('Template CSV diunduh.');
  }

  /* ---------------- Admin: cadangan & reset ---------------- */
  function downloadBlob(filename, blob) {
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
  }

  function exportBackup() {
    var payload = {
      app: 'SIMPATI',
      version: 1,
      exportedAt: new Date().toISOString(),
      info: DATASET,
      db: DB
    };
    downloadBlob('simpati-cadangan-' + todayISO() + '.json', new Blob([JSON.stringify(payload)], { type: 'application/json' }));
    addLog('Ekspor cadangan', nf.format(ASET.length) + ' aset · ' + (DATASET.source || ''));
    toast('Cadangan JSON berhasil diunduh.');
  }

  function readRestoreFile(file) {
    if (!file) return;
    if (!/\.json$/i.test(file.name)) { toast('File cadangan harus berformat .json.'); return; }
    var reader = new FileReader();
    reader.onload = function () {
      var payload;
      try { payload = JSON.parse(String(reader.result || '')); } catch (e) { toast('File cadangan tidak valid.'); return; }
      if (!payload || !payload.db || !Array.isArray(payload.db.aset) || !Array.isArray(payload.db.lokasi)) {
        toast('Struktur cadangan tidak dikenali.');
        return;
      }
      admin.pendingRestore = payload;
      var count = payload.db.aset.length;
      openModal(
        '<div class="modal-head"><div><h3>Pulihkan cadangan?</h3><p>Dataset aktif akan diganti dengan isi cadangan.</p></div>' +
        '<button class="icon-btn modal-close" data-close aria-label="Tutup"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="m6 6 12 12M18 6 6 18"/></svg></button></div>' +
        '<div class="modal-body"><div class="receipt">' +
          receiptRow('File', file.name) +
          receiptRow('Jumlah aset', nf.format(count) + ' unit') +
          receiptRow('Sumber', (payload.info && payload.info.source) || (payload.db.meta && payload.db.meta.source) || '—') +
          receiptRow('Diekspor', payload.exportedAt ? fmtDateTime(payload.exportedAt) : '—') +
        '</div><p class="modal-warn">Data aset dan tambahan manual saat ini akan diganti. Riwayat perpindahan yang asetnya tidak ada akan dihapus.</p></div>' +
        '<div class="modal-foot"><button class="btn btn-ghost" data-close>Batal</button><button class="btn btn-primary" data-confirm-restore>Pulihkan data</button></div>'
      );
      if ($('#restore-file')) $('#restore-file').value = '';
    };
    reader.readAsText(file);
  }

  function finishRestore() {
    var payload = admin.pendingRestore;
    if (!payload) return;
    var info = {
      kind: (payload.info && payload.info.kind) || 'upload',
      source: (payload.info && payload.info.source) || 'Cadangan',
      fileName: (payload.info && payload.info.fileName) || '',
      sheet: (payload.info && payload.info.sheet) || '',
      updatedAt: new Date().toISOString(),
      importedAt: new Date().toISOString()
    };
    withMeta(payload.db, info);
    applyDataset(payload.db, info);
    persistDataset();
    addLog('Pulihkan cadangan', nf.format(payload.db.aset.length) + ' aset · ' + info.source);
    admin.pendingRestore = null;
    closeModal();
    renderAdmin();
    toast('Cadangan berhasil dipulihkan.');
  }

  function confirmResetDataset() {
    openModal(
      '<div class="modal-head"><div><h3>Kembalikan data bawaan?</h3><p>Data unggahan dan barang tambahan akan hilang.</p></div>' +
      '<button class="icon-btn modal-close" data-close aria-label="Tutup"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8"><path d="m6 6 12 12M18 6 6 18"/></svg></button></div>' +
      '<div class="modal-body"><p class="modal-warn">Dataset aktif akan diganti dengan data awal LOKTAIL OKTOBER 2026 (' + nf.format(DEFAULT_DB.aset.length) + ' aset). Riwayat perpindahan yang asetnya tidak ada akan dihapus. Tindakan ini tidak dapat dibatalkan.</p></div>' +
      '<div class="modal-foot"><button class="btn btn-ghost" data-close>Batal</button>' +
      '<button class="btn btn-danger" data-confirm-dataset-reset>Ya, kembalikan data bawaan</button></div>'
    );
  }

  function finishResetDataset() {
    if (window.DatasetStore) window.DatasetStore.clear().catch(function () { /* abaikan */ });
    applyDataset(DEFAULT_DB, {
      kind: 'default',
      source: DEFAULT_DB.meta.source || 'LOKTAIL OKTOBER 2026',
      fileName: '',
      sheet: 'ALL',
      updatedAt: new Date().toISOString()
    });
    addLog('Reset dataset', 'Kembali ke data bawaan · ' + nf.format(DEFAULT_DB.aset.length) + ' aset');
    closeModal();
    renderAdmin();
    toast('Data bawaan berhasil dipulihkan.');
  }

  /* ---------------- Events ---------------- */
  function bind() {
    $$('.nav-item').forEach(function (b) {
      b.addEventListener('click', function () { go(b.getAttribute('data-view')); });
    });

    document.addEventListener('click', function (e) {
      var goBtn = e.target.closest('[data-goto]');
      if (goBtn) { closeModal(); go(goBtn.getAttribute('data-goto')); return; }

      var closeBtn = e.target.closest('[data-close]');
      if (closeBtn) { closeModal(); return; }

      var backdrop = e.target.closest('[data-backdrop]');
      if (backdrop && e.target === backdrop) { closeModal(); return; }

      var moveBtn = e.target.closest('[data-move]');
      if (moveBtn) {
        var ref = moveBtn.getAttribute('data-move');
        var a = ASET.find(function (x) { return x.ref === ref; });
        closeModal();
        go('perpindahan');
        if (a) { selectAsset(a); setTimeout(function () { $('#move-to').focus(); }, 120); }
        return;
      }

      var detailBtn = e.target.closest('[data-detail]');
      if (detailBtn) { openAssetDetail(detailBtn.getAttribute('data-detail')); return; }

      var delBtn = e.target.closest('[data-del]');
      if (delBtn) { confirmDelete(delBtn.getAttribute('data-del')); return; }

      var cdel = e.target.closest('[data-confirm-del]');
      if (cdel) { doDelete(cdel.getAttribute('data-confirm-del')); return; }

      if (e.target.closest('[data-lock-admin]')) { lockAdmin(); return; }

      var rpick = e.target.closest('[data-repair-pick]');
      if (rpick) {
        var rpAsset = ASET.find(function (x) { return x.ref === rpick.getAttribute('data-repair-pick'); });
        if (rpAsset) selectRepairAsset(rpAsset);
        return;
      }

      var rload = e.target.closest('[data-repair-load]');
      if (rload) { loadRepairRecord(rload.getAttribute('data-repair-load')); return; }

      var rdel = e.target.closest('[data-repair-del]');
      if (rdel) { confirmDeleteRepair(rdel.getAttribute('data-repair-del')); return; }

      var rcdel = e.target.closest('[data-confirm-repair-del]');
      if (rcdel) { doDeleteRepair(rcdel.getAttribute('data-confirm-repair-del')); return; }

      if (e.target.closest('[data-upload-remove]')) { uploadReset(); return; }
      if (e.target.closest('[data-upload-again]')) {
        uploadReset();
        if ($('#upload-file')) $('#upload-file').click();
        return;
      }
      if (e.target.closest('[data-confirm-import]')) { finishImport(); return; }
      if (e.target.closest('[data-confirm-restore]')) { finishRestore(); return; }
      if (e.target.closest('[data-confirm-dataset-reset]')) { finishResetDataset(); return; }

      if (e.target.closest('[data-confirm-reset]')) {
        state.transfers = [];
        state.overrides = {};
        save();
        closeModal();
        renderAll();
        resetMoveForm();
        if (admin.unlocked) addLog('Reset perpindahan', 'Semua catatan perpindahan dihapus');
        toast('Data perpindahan berhasil direset.');
        return;
      }

      var pickBtn = e.target.closest('[data-pick]');
      if (pickBtn) {
        var picked = ASET.find(function (x) { return x.ref === pickBtn.getAttribute('data-pick'); });
        if (picked) selectAsset(picked);
        return;
      }

      var depDetail = e.target.closest('[data-dep-detail]');
      if (depDetail) { openDepDetail(depDetail.getAttribute('data-dep-detail')); return; }

      var depPageBtn = e.target.closest('[data-dep-page]');
      if (depPageBtn && !depPageBtn.disabled) {
        var dp = +depPageBtn.getAttribute('data-dep-page');
        if (dp >= 1) { ui.depPage = dp; renderPenyusutan(); }
        return;
      }

      var pageBtn = e.target.closest('[data-page]');
      if (pageBtn && !pageBtn.disabled) {
        var p = +pageBtn.getAttribute('data-page');
        if (p >= 1) { ui.page = p; renderAset(); }
        return;
      }

      var clearBtn = e.target.closest('[data-clear]');
      if (clearBtn) {
        var t = clearBtn.getAttribute('data-clear');
        if (t === 'search') { ui.search = ''; $('#aset-search').value = ''; }
        if (t === 'sub') { ui.sub = ''; $('#aset-sub').value = ''; }
        if (t === 'loc') { ui.loc = ''; $('#aset-lokasi').value = ''; }
        if (t === 'dep-search') { ui.depSearch = ''; $('#dep-search').value = ''; }
        if (t === 'dep-kat') { ui.depKat = ''; $('#dep-kat').value = ''; }
        if (t === 'dep-status') { ui.depStatus = ''; $('#dep-status').value = ''; }
        if (t.indexOf('dep-') === 0) {
          ui.depPage = 1;
          renderPenyusutan();
        } else {
          ui.page = 1;
          renderAset();
        }
        return;
      }

      var depRow = e.target.closest('#dep-tbody tr[data-dep-ref]');
      if (depRow) { openDepDetail(depRow.getAttribute('data-dep-ref')); return; }

      var row = e.target.closest('#aset-tbody tr[data-ref]');
      if (row) { openAssetDetail(row.getAttribute('data-ref')); return; }

      var mrow = e.target.closest('.move-row[data-ref]');
      if (mrow) { openAssetDetail(mrow.getAttribute('data-ref')); return; }
    });

    document.addEventListener('keydown', function (e) {
      if (e.key === 'Escape') {
        if ($('#modal-root').innerHTML) closeModal();
        else if ($('#sidebar').classList.contains('is-open')) closeSidebar();
      }
    });

    $('#menu-btn').addEventListener('click', openSidebar);
    $('#overlay').addEventListener('click', closeSidebar);

    $('#aset-search').addEventListener('input', debounce(function (e) {
      ui.search = e.target.value; ui.page = 1; renderAset();
    }, 180));
    $('#aset-sub').addEventListener('change', function (e) { ui.sub = e.target.value; ui.page = 1; renderAset(); });
    $('#aset-lokasi').addEventListener('change', function (e) { ui.loc = e.target.value; ui.page = 1; renderAset(); });
    $('#aset-sort').addEventListener('change', function (e) { ui.sort = e.target.value; ui.page = 1; renderAset(); });
    $('#aset-pagesize').addEventListener('change', function (e) { ui.pageSize = +e.target.value; ui.page = 1; renderAset(); });

    $('#btn-export-aset').addEventListener('click', exportAssets);
    $('#btn-export-riwayat').addEventListener('click', exportHistory);

    $('#dep-search').addEventListener('input', debounce(function (e) {
      ui.depSearch = e.target.value; ui.depPage = 1; renderPenyusutan();
    }, 180));
    $('#dep-kat').addEventListener('change', function (e) { ui.depKat = e.target.value; ui.depPage = 1; renderPenyusutan(); });
    $('#dep-status').addEventListener('change', function (e) { ui.depStatus = e.target.value; ui.depPage = 1; renderPenyusutan(); });
    $('#dep-sort').addEventListener('change', function (e) { ui.depSort = e.target.value; ui.depPage = 1; renderPenyusutan(); });
    $('#dep-pagesize').addEventListener('change', function (e) { ui.depPageSize = +e.target.value; ui.depPage = 1; renderPenyusutan(); });
    $('#btn-export-penyusutan').addEventListener('click', exportDepreciation);

    /* picker */
    var picker = $('#asset-picker');
    picker.addEventListener('input', function () {
      if (ui.current && picker.value !== ui.current.name) {
        ui.current = null;
        $('#selected-asset').hidden = true;
        $('#move-from').value = '';
        updatePreview();
      }
      $('#picker-clear').hidden = !picker.value;
      renderPickerResults(picker.value);
    });
    picker.addEventListener('focus', function () {
      if (picker.value.trim().length >= 2) renderPickerResults(picker.value);
    });
    picker.addEventListener('keydown', function (e) {
      var opts = $$('#picker-results .picker-option');
      if (!opts.length) return;
      if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
        e.preventDefault();
        ui.pickerIndex = e.key === 'ArrowDown'
          ? Math.min(opts.length - 1, ui.pickerIndex + 1)
          : Math.max(0, ui.pickerIndex - 1);
        opts.forEach(function (o, i) { o.classList.toggle('is-focused', i === ui.pickerIndex); });
        opts[ui.pickerIndex].scrollIntoView({ block: 'nearest' });
      } else if (e.key === 'Enter') {
        e.preventDefault();
        if (ui.pickerIndex >= 0 && opts[ui.pickerIndex]) opts[ui.pickerIndex].click();
      }
    });
    $('#picker-clear').addEventListener('click', function () { clearAsset(); picker.focus(); });

    /* form */
    $('#move-date').value = todayISO();
    $('#move-form').addEventListener('submit', submitMove);
    $('#move-reset').addEventListener('click', function () { resetMoveForm(); });
    ['move-to', 'move-date', 'move-pic', 'move-reason'].forEach(function (id) {
      $('#' + id).addEventListener('input', updatePreview);
      $('#' + id).addEventListener('change', updatePreview);
    });

    $('#riwayat-search').addEventListener('input', debounce(function (e) {
      ui.histSearch = e.target.value; renderHistory();
    }, 180));
    $('#riwayat-reason').addEventListener('change', function (e) {
      ui.histReason = e.target.value; renderHistory();
    });

    /* tambah barang */
    $('#add-form').addEventListener('submit', submitAdd);
    ['add-ref', 'add-name', 'add-sub', 'add-loc', 'add-date', 'add-price', 'add-brand', 'add-kel', 'add-desc'].forEach(function (id) {
      $('#' + id).addEventListener('input', renderAddPreview);
      $('#' + id).addEventListener('change', renderAddPreview);
    });
    $('#add-reset').addEventListener('click', resetAddForm);

    /* perbaikan */
    var rPicker = $('#repair-picker');
    rPicker.addEventListener('input', function () {
      if (ui.repairAsset && rPicker.value !== ui.repairAsset.name) {
        ui.repairAsset = null;
        $('#repair-selected').hidden = true;
        $('#repair-price').value = '';
        $('#repair-estimate').value = '';
        renderRepairResult();
      }
      $('#repair-picker-clear').hidden = !rPicker.value;
      renderRepairPickerResults(rPicker.value);
    });
    rPicker.addEventListener('focus', function () {
      if (rPicker.value.trim().length >= 2) renderRepairPickerResults(rPicker.value);
    });
    rPicker.addEventListener('keydown', function (e) {
      var opts = $$('#repair-picker-results .picker-option');
      if (!opts.length) return;
      if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
        e.preventDefault();
        ui.repairPickerIndex = e.key === 'ArrowDown'
          ? Math.min(opts.length - 1, ui.repairPickerIndex + 1)
          : Math.max(0, ui.repairPickerIndex - 1);
        opts.forEach(function (o, i) { o.classList.toggle('is-focused', i === ui.repairPickerIndex); });
        opts[ui.repairPickerIndex].scrollIntoView({ block: 'nearest' });
      } else if (e.key === 'Enter') {
        e.preventDefault();
        if (ui.repairPickerIndex >= 0 && opts[ui.repairPickerIndex]) opts[ui.repairPickerIndex].click();
      }
    });
    $('#repair-picker-clear').addEventListener('click', function () { clearRepairAsset(); rPicker.focus(); });
    $('#repair-cost').addEventListener('input', renderRepairResult);
    $('#repair-form').addEventListener('submit', submitRepair);
    $('#repair-reset').addEventListener('click', resetRepairForm);
    $('#btn-export-repairs').addEventListener('click', exportRepairs);

    /* admin */
    $('#admin-login-form').addEventListener('submit', function (e) {
      e.preventDefault();
      var pin = $('#admin-pin').value.trim();
      var err = $('#admin-login-error');
      if (!pin) {
        err.hidden = false;
        err.textContent = 'PIN admin wajib diisi.';
        $('#admin-pin').classList.add('is-error');
        return;
      }
      if (!loginAdmin(pin)) {
        err.hidden = false;
        err.textContent = 'PIN salah. Coba lagi atau gunakan PIN awal 1234.';
        $('#admin-pin').classList.add('is-error');
        $('#admin-pin').select();
        return;
      }
      err.hidden = true;
      $('#admin-pin').classList.remove('is-error');
    });

    $('#admin-pin-form').addEventListener('submit', function (e) {
      e.preventDefault();
      var pin = $('#new-pin').value.trim();
      var err = $('#pin-error');
      if (!/^\d{4,12}$/.test(pin)) {
        err.hidden = false;
        err.textContent = 'PIN baru harus terdiri dari 4–12 digit angka.';
        return;
      }
      try { localStorage.setItem(ADMIN_PIN_KEY, pinHash(pin)); } catch (er) {
        err.hidden = false;
        err.textContent = 'PIN tidak dapat disimpan di peramban ini.';
        return;
      }
      err.hidden = true;
      $('#new-pin').value = '';
      addLog('Ganti PIN', 'PIN panel admin diperbarui');
      toast('PIN admin berhasil diperbarui.');
    });

    $('#btn-reset-dataset').addEventListener('click', confirmResetDataset);
    $('#btn-reset-moves-admin').addEventListener('click', confirmReset);
    $('#btn-backup').addEventListener('click', exportBackup);
    $('#restore-file').addEventListener('change', function (e) { readRestoreFile(e.target.files && e.target.files[0]); });
    $('#btn-template').addEventListener('click', downloadTemplate);
    $('#btn-clear-log').addEventListener('click', function () {
      admin.log = [];
      saveAdminLog();
      renderAdminLog();
      toast('Log aktivitas dibersihkan.');
    });

    /* unggah excel */
    $('#upload-file').addEventListener('change', function (e) { handleUploadFile(e.target.files && e.target.files[0]); });
    var dz = $('#dropzone');
    ['dragenter', 'dragover'].forEach(function (ev) {
      dz.addEventListener(ev, function (e) { e.preventDefault(); dz.classList.add('is-drag'); });
    });
    ['dragleave', 'drop'].forEach(function (ev) {
      dz.addEventListener(ev, function (e) { e.preventDefault(); dz.classList.remove('is-drag'); });
    });
    dz.addEventListener('drop', function (e) {
      var f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
      if (f) handleUploadFile(f);
    });

    $('#upload-sheet').addEventListener('change', function (e) {
      upload.sheetIndex = +e.target.value;
      upload.sheetName = upload.sheets[upload.sheetIndex].name;
      upload.rows = upload.sheets[upload.sheetIndex].rows;
      upload.headerRow = detectHeaderRow(upload.rows);
      applyUploadHeaders();
      upload.step = 2;
      renderUploadUI();
    });
    $('#btn-apply-header').addEventListener('click', function () {
      var n = +$('#upload-header').value;
      if (!n || n < 1) n = 1;
      if (n > upload.rows.length) n = upload.rows.length;
      upload.headerRow = n;
      applyUploadHeaders();
      upload.step = 2;
      renderUploadUI();
    });
    $('#btn-to-review').addEventListener('click', function () {
      if (!upload.built) return;
      upload.step = 3;
      renderUploadUI();
    });
    $('#btn-cancel-upload').addEventListener('click', uploadReset);
    $('#btn-do-import').addEventListener('click', confirmImport);
    $('#upload-mode').addEventListener('change', function (e) {
      if (e.target.name !== 'import-mode') return;
      upload.mode = e.target.value;
      renderUploadStats();
      renderUploadReview();
    });

    document.addEventListener('change', function (e) {
      var mapSel = e.target.closest('[data-map]');
      if (!mapSel) return;
      upload.map[mapSel.getAttribute('data-map')] = mapSel.value === '' ? '' : +mapSel.value;
      refreshUpload();
      renderUploadStats();
      renderUploadReview();
    });

    $('#btn-reset').addEventListener('click', confirmReset);
  }

  function debounce(fn, ms) {
    var t;
    return function (e) { clearTimeout(t); var ev = e; t = setTimeout(function () { fn(ev); }, ms); };
  }

  /* ---------------- Init ---------------- */
  function renderAll() {
    renderDashboard();
    renderAset();
    renderHistory();
    updatePreview();
    renderRepairs();
    renderPenyusutan();
    $('#nav-count-aset').textContent = nf.format(ASET.length);
    $('#nav-count-riwayat').textContent = state.transfers.length || '';
  }

  function startApp() {
    var hasState = load();
    if (!hasState) {
      seedExamples();
      save();
    } else {
      rebuildOverrides();
    }

    admin.unlocked = isAdminSession();
    loadAdminLog();

    $('#today-chip').textContent = dfLong.format(new Date());
    buildFilters();
    bind();
    renderAll();
    resetMoveForm();
    resetAddForm();
    resetRepairForm();
    updateSourceCard();
    renderAdmin();

    if (location.hash === '#perpindahan') go('perpindahan');
    if (location.hash === '#perbaikan') go('perbaikan');
    if (location.hash === '#penyusutan') go('penyusutan');
    if (location.hash === '#admin') go('admin');
    if (location.hash === '#tambah') go('tambah');
  }

  function init() {
    buildData(DEFAULT_DB);

    var saved = (window.DatasetStore && window.DatasetStore.supported())
      ? window.DatasetStore.get().catch(function () { return null; })
      : Promise.resolve(null);

    saved.then(function (rec) {
      if (rec && rec.db && Array.isArray(rec.db.aset) && rec.db.aset.length) {
        buildData(rec.db);
        if (rec.info) {
          DATASET = {
            kind: rec.info.kind || 'custom',
            source: rec.info.source || (rec.db.meta && rec.db.meta.source) || 'Data Aset',
            fileName: rec.info.fileName || '',
            sheet: rec.info.sheet || '',
            updatedAt: rec.info.updatedAt || rec.savedAt || '',
            importedAt: rec.info.importedAt || ''
          };
        }
      }
      startApp();
    }).catch(function () { startApp(); });
  }

  init();
})();
