var SPREADSHEET_ID = ''; // Kosongkan jika script terikat langsung (bound) ke Google Sheet
var SCAN_LIMIT_PER_DAY = 30; // maksimal scan struk per hari (ubah sesuai kebutuhan)
var DEFAULT_DOMPET = [
  { n: 'Cash',  i: '💵', logo: '' },
  { n: 'BCA',   i: '🏦', logo: '' },
  { n: 'Gopay', i: '📱', logo: '' },
  { n: 'Jago',  i: '🦁', logo: '' }
];
var KATEGORI_CFG = {
  keluar: [
    { n: 'Makan', i: '🍜', sub: [
      { n: 'Makanan', i: '🍛' },
      { n: 'Minuman', i: '🥤' },
      { n: 'Cafe', i: '☕' },
      { n: 'Snack Jajan', i: '🍿' },
      { n: 'Belanja Dapur', i: '🥬' }
    ]},
    { n: 'Transportasi', i: '🛵', sub: [
      { n: 'Bensin', i: '⛽' },
      { n: 'Parkir & Tol', i: '🅿️' },
      { n: 'Ojek / Taksi Online', i: '🚕' },
      { n: 'Servis Kendaraan', i: '🔧' }
    ]},
    { n: 'Tagihan', i: '🧾', sub: [
      { n: 'Listrik', i: '⚡' },
      { n: 'Air', i: '💧' },
      { n: 'Internet', i: '📶' },
      { n: 'Pulsa', i: '📱' }
    ]},
    { n: 'Produktivitas', i: '🎯', sub: [
      { n: 'Kursus', i: '🎓' },
      { n: 'Buku', i: '📚' },
      { n: 'AI Subs', i: '🤖' },
      { n: 'Apps Editor', i: '🎬' }
    ]},
    { n: 'Kesehatan', i: '🩺', sub: [
      { n: 'Obat & Vitamin', i: '💊' },
      { n: 'Dokter & Klinik', i: '🏥' },
      { n: 'Asuransi', i: '🛡️' },
      { n: 'BPJS', i: '💳' }
    ]},
    { n: 'Family & Lover', i: '❤️', sub: [
      { n: 'Keluarga', i: '👨👩👧👦' },
      { n: 'Lover', i: '💑' },
      { n: 'Friend', i: '🤝' }
    ]},
    { n: 'Belanja', i: '🛍️', sub: [
      { n: 'Pakaian', i: '👕' },
      { n: 'Elektronik & Gear', i: '📷' },
      { n: 'Kebutuhan Rumah', i: '🏠' }
    ]},
    { n: 'Olahraga & Hobi', i: '🏃', sub: [
      { n: 'Lari & Race', i: '👟' },
      { n: 'Hiburan', i: '🎮' },
      { n: 'Personal Care', i: '💆' }
    ]},
    { n: 'Lainnya', i: '📦', sub: [] }
  ],
  masuk: [
    { n: 'Gaji', i: '💼', sub: [{ n: 'Gaji Utama', i: '💰' }, { n: 'Bonus & THR', i: '🎁' }] },
    { n: 'Bisnis & Freelance', i: '🚀', sub: [{ n: 'Happy POY', i: '🧸' }, { n: 'Project Desain', i: '🎨' }] },
    { n: 'Investasi', i: '📈', sub: [{ n: 'Dividen', i: '💵' }, { n: 'Bunga & Cashback', i: '✨' }] },
    { n: 'Pemasukan Lain', i: '📥', sub: [] }
  ],
  transfer: [
    { n: 'Antar Dompet', i: '🔄', sub: [{ n: 'Pindah Saldo', i: '💸' }] },
    { n: 'Dana Darurat', i: '🛡️', sub: [{ n: 'Setor / Beli', i: '➕' }, { n: 'Cairkan / Tarik', i: '➖' }] },
    { n: 'Reksadana', i: '🌱', sub: [{ n: 'Setor / Beli', i: '➕' }, { n: 'Cairkan / Tarik', i: '➖' }] },
    { n: 'Saham', i: '📊', sub: [{ n: 'Setor / Beli', i: '➕' }, { n: 'Cairkan / Tarik', i: '➖' }] }
  ]
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Catatan Keuangan')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

function getSs_() {
  return SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

function findSheetByNames_(ss, candidates) {
  var sheets = ss.getSheets();
  for (var c = 0; c < candidates.length; c++) {
    var target = candidates[c].toLowerCase();
    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getName().toLowerCase() === target) return sheets[i];
    }
  }
  return null;
}

function ensureSheets_() {
  var ss = getSs_();
  var txSh = findSheetByNames_(ss, ['Transaksi', 'Transactions', 'Catatan']);
  if (!txSh) {
    txSh = ss.insertSheet('Transaksi');
    txSh.appendRow(['ID', 'Tanggal', 'Tipe', 'Kategori', 'SubKategori', 'Jumlah', 'Dompet', 'DompetTujuan', 'Catatan']);
    txSh.getRange(1, 1, 1, 9).setFontWeight('bold');
  }
  var budSh = findSheetByNames_(ss, ['Budget', 'Budgets', 'Anggaran']);
  if (!budSh) {
    budSh = ss.insertSheet('Budget');
    budSh.appendRow(['Bulan', 'Kategori', 'Jumlah']);
    budSh.getRange(1, 1, 1, 3).setFontWeight('bold');
  }
  var saSh = findSheetByNames_(ss, ['SaldoAwal', 'Saldo Awal', 'Opening']);
  if (!saSh) {
    saSh = ss.insertSheet('SaldoAwal');
    saSh.appendRow(['Bulan', 'Dompet', 'Jumlah']);
    saSh.getRange(1, 1, 1, 3).setFontWeight('bold');
  }
  var dompetSh = findSheetByNames_(ss, ['Dompet', 'Wallets', 'Akun']);
  if (!dompetSh) {
    dompetSh = ss.insertSheet('Dompet');
    dompetSh.appendRow(['Nama', 'Ikon', 'LogoData', 'Aktif']);
    dompetSh.getRange(1, 1, 1, 4).setFontWeight('bold');
  }
  if (dompetSh.getLastRow() <= 1) {
    DEFAULT_DOMPET.forEach(function(d) {
      dompetSh.appendRow([d.n, d.i, d.logo || '', 'YA']);
    });
  }
  var katSh = findSheetByNames_(ss, ['Kategori', 'Categories']);
  if (!katSh) {
    katSh = ss.insertSheet('Kategori');
    katSh.appendRow(['Tipe', 'Kategori', 'Ikon', 'SubKategori', 'IkonSub', 'Aktif']);
    katSh.getRange(1, 1, 1, 6).setFontWeight('bold');
  }
  if (katSh.getLastRow() <= 1) {
    var initRows = [];
    var tipeMap = { keluar: 'Pengeluaran', masuk: 'Pemasukan', transfer: 'Transfer' };
    Object.keys(KATEGORI_CFG).forEach(function(k) {
      var tpLabel = tipeMap[k];
      KATEGORI_CFG[k].forEach(function(cat) {
        if (cat.sub && cat.sub.length) {
          cat.sub.forEach(function(s) {
            initRows.push([tpLabel, cat.n, cat.i, s.n, s.i, 'YA']);
          });
        } else {
          initRows.push([tpLabel, cat.n, cat.i, '', '', 'YA']);
        }
      });
    });
    if (initRows.length) {
      katSh.getRange(2, 1, initRows.length, 6).setValues(initRows);
    }
  }
  var tgSh = findSheetByNames_(ss, ['Target', 'Targets']);
  if (!tgSh) {
    tgSh = ss.insertSheet('Target');
    tgSh.appendRow(['Kategori', 'Target']);
    tgSh.getRange(1, 1, 1, 2).setFontWeight('bold');
  }
  var evSh = findSheetByNames_(ss, ['Events', 'Event', 'Acara']);
  if (!evSh) {
    evSh = ss.insertSheet('Events');
    evSh.appendRow(['ID', 'Nama', 'Ikon', 'MataUang', 'Kurs', 'Budget', 'TglMulai', 'TglSelesai', 'Status', 'KecualiReport']);
    evSh.getRange(1, 1, 1, 10).setFontWeight('bold');
  }
  return { ss: ss, txSh: txSh, budSh: budSh, saSh: saSh, dompetSh: dompetSh, katSh: katSh, tgSh: tgSh, evSh: evSh };
}

function ensureTxEventCols_(sh) {
  var lastCol = Math.max(9, sh.getLastColumn());
  var raw = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var norm = raw.map(function(h) { return String(h || '').trim().toLowerCase().replace(/[^a-z0-9]/g, ''); });
  var needed = ['Event', 'MataUang', 'NominalAsing'];
  needed.forEach(function(colName) {
    var key = colName.toLowerCase();
    if (norm.indexOf(key) < 0) {
      var nextCol = sh.getLastColumn() + 1;
      sh.getRange(1, nextCol).setValue(colName).setFontWeight('bold');
      norm.push(key);
    }
  });
}

function mapHeaders_(sh, defaults) {
  var lastCol = Math.max(sh.getLastColumn(), Object.keys(defaults).length);
  var raw = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var headers = raw.map(function(h) { return String(h || '').trim().toLowerCase().replace(/[^a-z0-9]/g, ''); });
  var out = {};
  Object.keys(defaults).forEach(function(key) {
    var aliases = defaults[key].aliases;
    var idx = -1;
    for (var a = 0; a < aliases.length; a++) {
      var found = headers.indexOf(aliases[a]);
      if (found >= 0) { idx = found; break; }
    }
    if (idx < 0) {
      for (var b = 0; b < aliases.length && idx < 0; b++) {
        if (aliases[b].length < 4) continue;
        for (var h = 0; h < headers.length; h++) {
          if (headers[h].indexOf(aliases[b]) === 0) { idx = h; break; }
        }
      }
    }
    out[key] = idx >= 0 ? idx : defaults[key].fallback;
  });
  return out;
}

function txCols_(sh) {
  return mapHeaders_(sh, {
    id:       { aliases: ['id', 'txid'], fallback: 0 },
    date:     { aliases: ['tanggal', 'date', 'tgl'], fallback: 1 },
    type:     { aliases: ['tipe', 'type', 'jenis'], fallback: 2 },
    cat:      { aliases: ['kategori', 'category', 'cat'], fallback: 3 },
    sub:      { aliases: ['subkategori', 'subcategory', 'sub'], fallback: 4 },
    amount:   { aliases: ['jumlah', 'amount', 'nominal', 'nilai'], fallback: 5 },
    wallet:   { aliases: ['dompet', 'wallet', 'akun', 'sumber'], fallback: 6 },
    walletTo: { aliases: ['dompettujuan', 'walletto', 'tujuan'], fallback: 7 },
    note:     { aliases: ['catatan', 'note', 'keterangan', 'deskripsi'], fallback: 8 },
    eventId:  { aliases: ['event', 'eventid', 'acara'], fallback: 9 },
    currency: { aliases: ['matauang', 'currency', 'valas'], fallback: 10 },
    fxAmount: { aliases: ['nominalasing', 'fxamount', 'nilaiasing'], fallback: 11 }
  });
}

function slug_(s) { return String(s || '').toLowerCase().replace(/[^a-z0-9]/g, ''); }

function readKategoriCfg_(katSh) {
  var last = katSh ? katSh.getLastRow() : 0;
  if (last <= 1) return KATEGORI_CFG;
  var rows = katSh.getRange(2, 1, last - 1, 6).getValues();
  var out = { keluar: [], masuk: [], transfer: [] };
  var map = { keluar: {}, masuk: {}, transfer: {} };
  rows.forEach(function(r) {
    var tpRaw = String(r[0] || '').trim().toLowerCase();
    var catName = String(r[1] || '').trim();
    var catIcon = String(r[2] || '📦').trim() || '📦';
    var subName = String(r[3] || '').trim();
    var subIcon = String(r[4] || '•').trim() || '•';
    var aktif = String(r[5] || 'YA').trim().toUpperCase();
    if (!catName || aktif === 'TIDAK') return;
    var key = tpRaw === 'pemasukan' ? 'masuk' : (tpRaw === 'transfer' ? 'transfer' : 'keluar');
    if (!map[key][catName]) {
      var obj = { n: catName, i: catIcon, sub: [] };
      map[key][catName] = obj;
      out[key].push(obj);
    }
    if (subName) {
      var exists = map[key][catName].sub.some(function(s) { return s.n.toLowerCase() === subName.toLowerCase(); });
      if (!exists) {
        map[key][catName].sub.push({ n: subName, i: subIcon });
      }
    }
  });
  return (out.keluar.length || out.masuk.length) ? out : KATEGORI_CFG;
}

function canonCat_(tp, cat, katCfg) {
  var cfg = katCfg || KATEGORI_CFG;
  var key = tp === 'Pemasukan' ? 'masuk' : (tp === 'Transfer' ? 'transfer' : 'keluar');
  var list = cfg[key] || [], k = slug_(cat);
  if (!k) return cat;
  for (var i = 0; i < list.length; i++) if (slug_(list[i].n) === k) return list[i].n;
  if (k.length >= 4) {
    for (var j = 0; j < list.length; j++) {
      var n = slug_(list[j].n);
      if (n.indexOf(k) === 0 || k.indexOf(n) === 0) return list[j].n;
    }
  }
  return cat;
}

function canonSub_(tp, cat, sub, katCfg) {
  var cfg = katCfg || KATEGORI_CFG;
  var key = tp === 'Pemasukan' ? 'masuk' : (tp === 'Transfer' ? 'transfer' : 'keluar');
  var c = (cfg[key] || []).filter(function(x) { return x.n === cat; })[0];
  var k = slug_(sub);
  if (!c || !k) return sub;
  for (var i = 0; i < c.sub.length; i++) if (slug_(c.sub[i].n) === k) return c.sub[i].n;
  return sub;
}

function findTxRow_(sh, col, id) {
  id = String(id).trim();
  if (/^ROW-\d+$/.test(id)) {
    var r = Number(id.slice(4));
    return (r >= 2 && r <= sh.getLastRow()) ? r : -1;
  }
  var last = sh.getLastRow();
  if (last <= 1) return -1;
  var ids = sh.getRange(2, col.id + 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === id) return i + 2;
  }
  return -1;
}

function readDompetList_(dompetSh) {
  var last = dompetSh.getLastRow();
  if (last <= 1) return DEFAULT_DOMPET.slice();
  var numRows = last - 1;
  var rows = dompetSh.getRange(2, 1, numRows, 4).getValues();
  var out = [];
  rows.forEach(function(r) {
    var nama = String(r[0] || '').trim();
    var aktif = String(r[3] || 'YA').trim().toUpperCase();
    if (nama && aktif !== 'TIDAK') {
      out.push({
        n: nama,
        i: String(r[1] || '💳').trim() || '💳',
        logo: String(r[2] || '').trim()
      });
    }
  });
  return out.length ? out : DEFAULT_DOMPET.slice();
}

function readArchivedSet_(dompetSh) {
  var set = {}, last = dompetSh.getLastRow();
  if (last <= 1) return set;
  dompetSh.getRange(2, 1, last - 1, 4).getValues().forEach(function(r) {
    var nm = String(r[0] || '').trim();
    if (nm && String(r[3] || 'YA').trim().toUpperCase() === 'TIDAK') set[nm] = true;
  });
  return set;
}

function fmtDate_(val) {
  if (!val) return '';
  if (Object.prototype.toString.call(val) === '[object Date]' && !isNaN(val.getTime())) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  var s = String(val).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
  var d = new Date(s);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return s;
}

function normYm_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM');
  }
  var s = String(v || '').trim();
  if (/^\d{4}-\d{2}/.test(s)) return s.slice(0, 7);
  var d = new Date(s);
  if (s && !isNaN(d.getTime())) return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM');
  return s;
}

function prevYm_(ym, back) {
  var y = Number(ym.slice(0, 4));
  var m = Number(ym.slice(5, 7)) - 1 - back;
  y += Math.floor(m / 12);
  m = ((m % 12) + 12) % 12;
  return y + '-' + ('0' + (m + 1)).slice(-2);
}

function isWithdraw_(sub) {
  return /cair|tarik/i.test(String(sub || ''));
}

function parseAmount_(v) {
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  if (v === null || v === undefined || v === '') return 0;
  var s = String(v).trim();
  var neg = /^\(.*\)$/.test(s) || /^-/.test(s.replace(/^[^\d(-]+/, ''));
  s = s.replace(/[^\d.,]/g, '');
  if (!s) return 0;
  var lc = s.lastIndexOf(','), ld = s.lastIndexOf('.');
  if (lc >= 0 && ld >= 0) {
    if (lc > ld) s = s.replace(/\./g, '').replace(',', '.');
    else s = s.replace(/,/g, '');
  } else if (lc >= 0) {
    s = /^\d{1,3}(,\d{3})+$/.test(s) ? s.replace(/,/g, '') : s.replace(',', '.');
  } else if (ld >= 0) {
    if (/^\d{1,3}(\.\d{3})+$/.test(s) || (s.match(/\./g) || []).length > 1) s = s.replace(/\./g, '');
  }
  var n = Number(s);
  if (isNaN(n)) return 0;
  return neg ? -n : n;
}

/* ===== SERVER CACHE HELPER (Optimasi Kecepatan getData) ===== */
function getDataCacheKey_(ym) {
  var ver = PropertiesService.getScriptProperties().getProperty('DATA_CACHE_VER') || '1';
  return 'GETDATA_V' + ver + '_' + ym;
}

function invalidateDataCache_() {
  var props = PropertiesService.getScriptProperties();
  var ver = (Number(props.getProperty('DATA_CACHE_VER')) || 1) + 1;
  props.setProperty('DATA_CACHE_VER', String(ver));
}

function getData(ym) {
  var tz = Session.getScriptTimeZone();
  var today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var targetYm = (ym && /^\d{4}-\d{2}$/.test(ym)) ? ym : today.slice(0, 7);

  var cache = CacheService.getScriptCache();
  var cacheKey = getDataCacheKey_(targetYm);
  var cachedStr = cache.get(cacheKey);
  if (cachedStr) {
    try {
      var cachedObj = JSON.parse(cachedStr);
      cachedObj.today = today;
      return cachedObj;
    } catch (ignore) {}
  }

  var env = ensureSheets_();
  var dompetCfg = readDompetList_(env.dompetSh);
  var archivedSet = readArchivedSet_(env.dompetSh);
  var katCfg = readKategoriCfg_(env.katSh);
  var dompetMap = {};
  dompetCfg.forEach(function(d) {
    dompetMap[d.n] = {
      name: d.n,
      icon: d.i,
      logo: d.logo || '',
      opening: 0,
      monthIn: 0,
      monthOut: 0,
      balance: 0,
      baseMonth: '0000-00'
    };
  });

  // 1. Baca Saldo Awal
  var saRows = env.saSh.getLastRow() > 1
    ? env.saSh.getRange(2, 1, env.saSh.getLastRow() - 1, Math.max(3, env.saSh.getLastColumn())).getValues()
    : [];
  var openingByMonth = {};
  saRows.forEach(function(r) {
    var b = normYm_(r[0]);
    var w = String(r[1] || '').trim();
    var amt = parseAmount_(r[2]);
    if (!b || !w) return;
    if (!openingByMonth[b]) openingByMonth[b] = {};
    openingByMonth[b][w] = amt;
  });
  dompetCfg.forEach(function(d) {
    var openVal = 0;
    var baseM = '0000-00';
    if (openingByMonth[targetYm] && openingByMonth[targetYm][d.n] !== undefined) {
      openVal = openingByMonth[targetYm][d.n];
      baseM = targetYm;
    } else {
      var prevKeys = Object.keys(openingByMonth).filter(function(k) {
        return k < targetYm && openingByMonth[k][d.n] !== undefined;
      }).sort();
      if (prevKeys.length) {
        var lastKey = prevKeys[prevKeys.length - 1];
        openVal = openingByMonth[lastKey][d.n];
        baseM = lastKey;
      }
    }
    dompetMap[d.n].opening = openVal;
    dompetMap[d.n].balance = openVal;
    dompetMap[d.n].baseMonth = baseM;
  });

  // 2. Baca Budget (Auto Carry-Over dari bulan sebelumnya)
  var bRows = env.budSh.getLastRow() > 1 ? env.budSh.getRange(2, 1, env.budSh.getLastRow() - 1, Math.max(3, env.budSh.getLastColumn())).getValues() : [];
  var budgetByCat = {};
  bRows.forEach(function(r) {
    var b = normYm_(r[0]);
    var cat = canonCat_('Pengeluaran', String(r[1] || '').trim(), katCfg);
    var amt = parseAmount_(r[2]);
    if (!b || !cat || !/^\d{4}-\d{2}$/.test(b)) return;
    if (!budgetByCat[cat]) budgetByCat[cat] = {};
    budgetByCat[cat][b] = amt;
  });
  var budgets = {};
  Object.keys(budgetByCat).forEach(function(cat) {
    var months = Object.keys(budgetByCat[cat]).filter(function(m) { return m <= targetYm; }).sort();
    if (months.length) {
      var latestMonth = months[months.length - 1];
      var val = budgetByCat[cat][latestMonth];
      if (val > 0) budgets[cat] = val;
    }
  });

  // 2b. Baca Target tabungan (per aset)
  var targets = { 'Dana Darurat': 0, 'Reksadana': 0, 'Saham': 0 };
  var tgLast = env.tgSh.getLastRow();
  if (tgLast > 1) {
    env.tgSh.getRange(2, 1, tgLast - 1, 2).getValues().forEach(function(r) {
      var k = String(r[0] || '').trim();
      if (targets.hasOwnProperty(k)) targets[k] = parseAmount_(r[1]);
    });
  }

  // 2c. Baca Events
  var eventsList = [], eventMap = {};
  var evLast = env.evSh.getLastRow();
  if (evLast > 1) {
    env.evSh.getRange(2, 1, evLast - 1, 10).getValues().forEach(function(r) {
      var evId = String(r[0] || '').trim();
      var evName = String(r[1] || '').trim();
      if (!evId || !evName) return;
      var evObj = {
        id: evId,
        name: evName,
        icon: String(r[2] || '✈️').trim() || '✈️',
        currency: String(r[3] || 'IDR').trim().toUpperCase() || 'IDR',
        rate: Number(r[4]) > 0 ? Number(r[4]) : 1,
        budget: parseAmount_(r[5]),
        startDate: fmtDate_(r[6]),
        endDate: fmtDate_(r[7]),
        status: String(r[8] || 'AKTIF').trim().toUpperCase() === 'SELESAI' ? 'SELESAI' : 'AKTIF',
        excludeReport: String(r[9] || 'TIDAK').trim().toUpperCase() === 'YA',
        spent: 0, income: 0, fxSpent: 0, fxIncome: 0, txCount: 0
      };
      eventMap[evId] = evObj;
      eventsList.push(evObj);
    });
  }

  // 3. Baca Transaksi
  var tCol = txCols_(env.txSh);
  var txLastRow = env.txSh.getLastRow();
  var txColsCount = Math.max(9, env.txSh.getLastColumn());
  var tRows = txLastRow > 1 ? env.txSh.getRange(2, 1, txLastRow - 1, txColsCount).getValues() : [];
  var txs = [];
  var income = 0, expense = 0;
  var catMap = {};
  var assetCats = ['Dana Darurat', 'Reksadana', 'Saham'];
  var assetTotal = { 'Dana Darurat': 0, 'Reksadana': 0, 'Saham': 0 };
  var assetMonth = { 'Dana Darurat': 0, 'Reksadana': 0, 'Saham': 0 };
  var prev3 = [prevYm_(targetYm, 1), prevYm_(targetYm, 2), prevYm_(targetYm, 3)];
  var prevDaily = {};
  prev3.forEach(function(p) { prevDaily[p] = {}; });
  var curDaily = {};

  tRows.forEach(function(r, rowIdx) {
    var ds = fmtDate_(r[tCol.date]);
    if (!ds || ds.length < 7) return;
    var id = String(r[tCol.id] || ('ROW-' + (rowIdx + 2))).trim();
    var rowYm = ds.slice(0, 7);
    var tp = String(r[tCol.type] || 'Pengeluaran').trim();
    var cat = canonCat_(tp, String(r[tCol.cat] || '').trim(), katCfg);
    var sub = canonSub_(tp, cat, String(r[tCol.sub] || '').trim(), katCfg);
    var amt = parseAmount_(r[tCol.amount]);
    var w = String(r[tCol.wallet] || 'Cash').trim() || 'Cash';
    var wTo = String(r[tCol.walletTo] || '').trim();
    var note = String(r[tCol.note] || '').trim();
    if (!amt && !cat) return;

    var evId = String(r[tCol.eventId] || '').trim();
    var curCode = String(r[tCol.currency] || '').trim().toUpperCase();
    var fxAmt = parseAmount_(r[tCol.fxAmount]);
    if (evId && eventMap[evId]) {
      eventMap[evId].txCount++;
      if (tp === 'Pengeluaran') { eventMap[evId].spent += amt; eventMap[evId].fxSpent += fxAmt; }
      else if (tp === 'Pemasukan') { eventMap[evId].income += amt; eventMap[evId].fxIncome += fxAmt; }
    }
    var skipAvg = !!(evId && eventMap[evId] && eventMap[evId].excludeReport);

    if (!dompetMap[w]) {
      dompetMap[w] = { name: w, icon: '💳', logo: '', opening: 0, monthIn: 0, monthOut: 0, balance: 0, baseMonth: '0000-00', archived: !!archivedSet[w] };
    }
    if (wTo && !dompetMap[wTo]) {
      dompetMap[wTo] = { name: wTo, icon: '💳', logo: '', opening: 0, monthIn: 0, monthOut: 0, balance: 0, baseMonth: '0000-00', archived: !!archivedSet[wTo] };
    }

    // Roll-over saldo bulan-bulan sebelumnya
    if (rowYm < targetYm) {
      var baseW = dompetMap[w].baseMonth || '0000-00';
      if (rowYm >= baseW) {
        if (tp === 'Pemasukan') {
          dompetMap[w].opening += amt;
          dompetMap[w].balance += amt;
        } else if (tp === 'Pengeluaran') {
          dompetMap[w].opening -= amt;
          dompetMap[w].balance -= amt;
        } else if (tp === 'Transfer') {
          if (cat === 'Antar Dompet' || !isWithdraw_(sub)) {
            dompetMap[w].opening -= amt;
            dompetMap[w].balance -= amt;
          } else if (isWithdraw_(sub)) {
            dompetMap[w].opening += amt;
            dompetMap[w].balance += amt;
          }
        }
      }
      if (tp === 'Transfer' && cat === 'Antar Dompet' && wTo) {
        var baseWTo = dompetMap[wTo].baseMonth || '0000-00';
        if (rowYm >= baseWTo) {
          dompetMap[wTo].opening += amt;
          dompetMap[wTo].balance += amt;
        }
      }
    }

    if (rowYm <= targetYm && tp === 'Transfer' && assetCats.indexOf(cat) >= 0) {
      var deltaAsset = isWithdraw_(sub) ? (0 - amt) : amt;
      assetTotal[cat] += deltaAsset;
      if (rowYm === targetYm) assetMonth[cat] += deltaAsset;
    }

    if (!skipAvg && prev3.indexOf(rowYm) >= 0 && tp === 'Pengeluaran') {
      var dayNum = Number(ds.slice(8, 10)) || 1;
      prevDaily[rowYm][dayNum] = (prevDaily[rowYm][dayNum] || 0) + amt;
    }

    if (rowYm === targetYm) {
      txs.push({
        id: id, date: ds, type: tp, cat: cat, sub: sub || null,
        amount: amt, wallet: w, walletTo: wTo || null, note: note,
        eventId: evId, currency: curCode, fxAmount: fxAmt
      });
      if (tp === 'Pemasukan') {
        income += amt;
        dompetMap[w].monthIn += amt;
        dompetMap[w].balance += amt;
      } else if (tp === 'Pengeluaran') {
        expense += amt;
        dompetMap[w].monthOut += amt;
        dompetMap[w].balance -= amt;
        var dIdx = Number(ds.slice(8, 10)) || 1;
        if (!skipAvg) curDaily[dIdx] = (curDaily[dIdx] || 0) + amt;
        if (!catMap[cat]) catMap[cat] = { cat: cat, amount: 0, subMap: {} };
        catMap[cat].amount += amt;
        var sKey = sub || 'Umum';
        catMap[cat].subMap[sKey] = (catMap[cat].subMap[sKey] || 0) + amt;
      } else if (tp === 'Transfer') {
        if (cat === 'Antar Dompet') {
          dompetMap[w].monthOut += amt;
          dompetMap[w].balance -= amt;
          if (wTo) {
            dompetMap[wTo].monthIn += amt;
            dompetMap[wTo].balance += amt;
          }
        } else if (isWithdraw_(sub)) {
          dompetMap[w].monthIn += amt;
          dompetMap[w].balance += amt;
        } else {
          dompetMap[w].monthOut += amt;
          dompetMap[w].balance -= amt;
        }
      }
    }
  });

  txs.sort(function(a, b) {
    if (a.date !== b.date) return a.date < b.date ? 1 : -1;
    return a.id < b.id ? 1 : -1;
  });

  var categories = Object.keys(catMap).map(function(k) {
    var c = catMap[k];
    var subs = Object.keys(c.subMap).map(function(sk) {
      return { sub: sk, amount: c.subMap[sk] };
    }).sort(function(a, b) {
      return b.amount - a.amount;
    });
    return { cat: c.cat, amount: c.amount, subs: subs };
  }).sort(function(a, b) {
    return b.amount - a.amount;
  });

  var wallets = Object.keys(dompetMap).map(function(k) { return dompetMap[k]; })
    .filter(function(w) { return !w.archived; });
  var totalWalletBalance = wallets.reduce(function(acc, w) { return acc + w.balance; }, 0);
  var totalAsset = assetTotal['Dana Darurat'] + assetTotal['Reksadana'] + assetTotal['Saham'];
  var saving = income - expense;

  // Smart Budget Alert (>= 80% Waspada, >= 100% Over)
  var budgetAlerts = [];
  Object.keys(budgets).forEach(function(bCat) {
    var limit = budgets[bCat];
    var used = catMap[bCat] ? catMap[bCat].amount : 0;
    if (limit > 0) {
      var pct = Math.round((used / limit) * 100);
      if (pct >= 80) {
        budgetAlerts.push({
          cat: bCat,
          limit: limit,
          used: used,
          pct: pct,
          status: pct >= 100 ? 'OVER' : 'WASPADA',
          diff: used - limit
        });
      }
    }
  });
  budgetAlerts.sort(function(a, b) { return b.pct - a.pct; });

  var yNum = Number(targetYm.slice(0, 4)), mNum = Number(targetYm.slice(5, 7));
  var daysInMonth = new Date(yNum, mNum, 0).getDate();
  var maxCurDay = targetYm === today.slice(0, 7) ? Math.min(daysInMonth, Number(today.slice(8, 10))) : daysInMonth;
  var currentArr = [], runCur = 0;
  for (var d = 1; d <= maxCurDay; d++) {
    runCur += (curDaily[d] || 0);
    currentArr.push(runCur);
  }
  var activePrevMonths = prev3.filter(function(p) { return Object.keys(prevDaily[p]).length > 0; });
  var avgArr = [];
  if (activePrevMonths.length) {
    var runPrev = {};
    activePrevMonths.forEach(function(p) { runPrev[p] = 0; });
    for (var day = 1; day <= daysInMonth; day++) {
      var sumDay = 0;
      activePrevMonths.forEach(function(p) {
        runPrev[p] += (prevDaily[p][day] || 0);
        sumDay += runPrev[p];
      });
      avgArr.push(Math.round(sumDay / activePrevMonths.length));
    }
  }

  var result = {
    ym: targetYm,
    today: today,
    income: income,
    expense: expense,
    saving: saving,
    rate: income > 0 ? Math.max(0, saving / income) : 0,
    freeCash: totalWalletBalance,
    totalAsset: totalAsset,
    totalSaving: totalWalletBalance + totalAsset,
    portfolio: [
      { cat: 'Dana Darurat', icon: '🛡️', total: assetTotal['Dana Darurat'], monthNet: assetMonth['Dana Darurat'], target: targets['Dana Darurat'] },
      { cat: 'Reksadana',    icon: '🌱', total: assetTotal['Reksadana'],    monthNet: assetMonth['Reksadana'],    target: targets['Reksadana'] },
      { cat: 'Saham',        icon: '📊', total: assetTotal['Saham'],        monthNet: assetMonth['Saham'],        target: targets['Saham'] }
    ],
    categories: categories,
    transactions: txs,
    budgets: budgets,
    budgetAlerts: budgetAlerts,
    wallets: wallets,
    dompet: dompetCfg,
    kategori: katCfg,
    events: eventsList,
    report: {
      days: daysInMonth,
      cur: currentArr,
      avg: avgArr.slice(0, currentArr.length),
      curTotal: expense,
      avgTotal: avgArr.length ? avgArr[currentArr.length - 1] : 0,
      avgMonths: activePrevMonths.length
    }
  };

  try {
    cache.put(cacheKey, JSON.stringify(result), 600);
  } catch (ignore) {}

  return result;
}

function searchTransactions(opts) {
  opts = opts || {};
  var query = opts.query, filterType = opts.type, filterWallet = opts.wallet;
  var env = ensureSheets_();
  var katCfg = readKategoriCfg_(env.katSh);
  var tCol = txCols_(env.txSh);
  var txLastRow = env.txSh.getLastRow();
  if (txLastRow <= 1) return [];
  var txColsCount = Math.max(9, env.txSh.getLastColumn());
  var tRows = env.txSh.getRange(2, 1, txLastRow - 1, txColsCount).getValues();
  var q = String(query || '').trim().toLowerCase();
  var fType = String(filterType || 'Semua').trim();
  var fWallet = String(filterWallet || 'Semua').trim();
  var fEvent = String(opts.eventId || 'Semua').trim();
  var results = [];
  for (var i = 0; i < tRows.length; i++) {
    var r = tRows[i];
    var ds = fmtDate_(r[tCol.date]);
    if (!ds || ds.length < 7) continue;
    var id = String(r[tCol.id] || ('ROW-' + (i + 2))).trim();
    var tp = String(r[tCol.type] || 'Pengeluaran').trim();
    var cat = canonCat_(tp, String(r[tCol.cat] || '').trim(), katCfg);
    var sub = canonSub_(tp, cat, String(r[tCol.sub] || '').trim(), katCfg);
    var amt = parseAmount_(r[tCol.amount]);
    var w = String(r[tCol.wallet] || 'Cash').trim() || 'Cash';
    var wTo = String(r[tCol.walletTo] || '').trim();
    var note = String(r[tCol.note] || '').trim();
    var evId = String(r[tCol.eventId] || '').trim();
    if (!amt && !cat) continue;
    if (fType !== 'Semua' && tp !== fType) continue;
    if (fEvent !== 'Semua') {
      if (fEvent === '__NONE__' && evId) continue;
      if (fEvent !== '__NONE__' && evId !== fEvent) continue;
    }
    if (fWallet !== 'Semua') {
      var matchFrom = (w === fWallet);
      var matchTo = (tp === 'Transfer' && cat === 'Antar Dompet' && wTo === fWallet);
      if (!matchFrom && !matchTo) continue;
    }
    if (q) {
      var hay = (cat + ' ' + sub + ' ' + note + ' ' + w + ' ' + wTo + ' ' + ds).toLowerCase();
      if (hay.indexOf(q) < 0) continue;
    }
    results.push({
      id: id, date: ds, type: tp, cat: cat, sub: sub || null,
      amount: amt, wallet: w, walletTo: wTo || null, note: note,
      eventId: String(r[tCol.eventId] || '').trim(),
      currency: String(r[tCol.currency] || '').trim().toUpperCase(),
      fxAmount: parseAmount_(r[tCol.fxAmount])
    });
  }
  results.sort(function(a, b) {
    if (a.date !== b.date) return a.date < b.date ? 1 : -1;
    return a.id < b.id ? 1 : -1;
  });
  return results.slice(0, 60);
}

function addWallet(payload, ym) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var env = ensureSheets_();
    var name = String((payload && payload.name) || '').trim();
    if (!name) throw new Error('Nama dompet tidak boleh kosong');
    if (name.length > 28) throw new Error('Nama dompet maksimal 28 karakter');
    var icon = String((payload && payload.icon) || '💳').trim() || '💳';
    var logoData = String((payload && payload.logoData) || '').trim();
    if (logoData && logoData.length > 48000) {
      throw new Error('Ukuran gambar terlalu besar untuk sel Google Sheet');
    }
    var sh = env.dompetSh;
    var last = sh.getLastRow();
    if (last > 1) {
      var numRows = last - 1;
      var existing = sh.getRange(2, 1, numRows, 4).getValues();
      for (var i = 0; i < existing.length; i++) {
        var exName = String(existing[i][0] || '').trim();
        if (exName.toLowerCase() === name.toLowerCase()) {
          var isAktif = String(existing[i][3] || 'YA').trim().toUpperCase();
          if (isAktif === 'TIDAK') {
            sh.getRange(i + 2, 1, 1, 4).setValues([[name, icon, logoData, 'YA']]);
            invalidateDataCache_();
            var initReOpen = Number(payload && payload.opening) || 0;
            if (initReOpen !== 0 && ym) return setOpening(name, initReOpen, ym);
            return getData(ym || null);
          }
          throw new Error('Dompet "' + exName + '" sudah ada');
        }
      }
    }
    sh.appendRow([name, icon, logoData, 'YA']);
    invalidateDataCache_();
    var opening = Number(payload && payload.opening) || 0;
    if (opening !== 0 && ym) {
      return setOpening(name, opening, ym);
    }
    return getData(ym || null);
  } finally {
    lock.releaseLock();
  }
}

function deleteWallet(name, ym) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var env = ensureSheets_();
    var sh = env.dompetSh;
    name = String(name || '').trim();
    if (!name) throw new Error('Nama dompet tidak valid');
    if (readDompetList_(sh).length <= 1) throw new Error('Minimal harus ada satu dompet aktif');
    var last = sh.getLastRow();
    if (last > 1) {
      var rows = sh.getRange(2, 1, last - 1, 4).getValues();
      for (var i = 0; i < rows.length; i++) {
        if (String(rows[i][0]).trim().toLowerCase() === name.toLowerCase()) {
          sh.getRange(i + 2, 4).setValue('TIDAK');
          invalidateDataCache_();
          return getData(ym || null);
        }
      }
    }
    throw new Error('Dompet "' + name + '" tidak ditemukan');
  } finally {
    lock.releaseLock();
  }
}

function addCategory(payload, ym) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var env = ensureSheets_();
    var sh = env.katSh;
    var type = String((payload && payload.type) || 'Pengeluaran').trim();
    var cat = String((payload && payload.cat) || '').trim();
    var icon = String((payload && payload.icon) || '📦').trim() || '📦';
    var sub = String((payload && payload.sub) || '').trim();
    var subIcon = String((payload && payload.subIcon) || '•').trim() || '•';
    if (!cat) throw new Error('Nama kategori tidak boleh kosong');
    if (cat.length > 28) throw new Error('Nama kategori maksimal 28 karakter');
    var last = sh.getLastRow();
    if (last > 1) {
      var rows = sh.getRange(2, 1, last - 1, 6).getValues();
      for (var i = 0; i < rows.length; i++) {
        var rType = String(rows[i][0] || '').trim().toLowerCase();
        var rCat = String(rows[i][1] || '').trim().toLowerCase();
        var rSub = String(rows[i][3] || '').trim().toLowerCase();
        if (rType === type.toLowerCase() && rCat === cat.toLowerCase() && rSub === sub.toLowerCase()) {
          if (String(rows[i][5] || '').trim().toUpperCase() === 'TIDAK') {
            sh.getRange(i + 2, 1, 1, 6).setValues([[type, cat, icon, sub, sub ? subIcon : '', 'YA']]);
            invalidateDataCache_();
            return getData(ym || null);
          }
          throw new Error(sub ? 'Sub-kategori "' + sub + '" sudah ada di ' + cat : 'Kategori "' + cat + '" sudah ada');
        }
      }
    }
    var newRow = [type, cat, icon, sub, sub ? subIcon : '', 'YA'];
    var insertAt = -1;
    var tailNames = ['lainnya', 'pemasukan lain'];
    if (last > 1 && tailNames.indexOf(cat.toLowerCase()) < 0) {
      var allRows = sh.getRange(2, 1, last - 1, 6).getValues();
      for (var j = 0; j < allRows.length; j++) {
        var rt = String(allRows[j][0] || '').trim().toLowerCase();
        var rc = String(allRows[j][1] || '').trim().toLowerCase();
        var ra = String(allRows[j][5] || 'YA').trim().toUpperCase();
        if (rt === type.toLowerCase() && ra !== 'TIDAK' && tailNames.indexOf(rc) >= 0) {
          insertAt = j + 2;
          break;
        }
      }
    }
    if (insertAt >= 2) {
      sh.insertRowBefore(insertAt);
      sh.getRange(insertAt, 1, 1, 6).setValues([newRow]);
    } else {
      sh.appendRow(newRow);
    }
    invalidateDataCache_();
    return getData(ym || null);
  } finally {
    lock.releaseLock();
  }
}

function reorderCategory(type, cat, dir, ym) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var env = ensureSheets_();
    var sh = env.katSh;
    var last = sh.getLastRow();
    if (last <= 1) return getData(ym || null);
    var rows = sh.getRange(2, 1, last - 1, 6).getValues();
    var tLow = String(type || 'Pengeluaran').trim().toLowerCase();
    var cLow = String(cat || '').trim().toLowerCase();

    var otherRows = [];
    var groupOrder = [];
    var groupMap = {};
    rows.forEach(function(r) {
      var rt = String(r[0] || '').trim().toLowerCase();
      var rc = String(r[1] || '').trim();
      var rcLow = rc.toLowerCase();
      var ra = String(r[5] || 'YA').trim().toUpperCase();
      if (rt === tLow && rc && ra !== 'TIDAK') {
        if (!groupMap[rcLow]) {
          groupMap[rcLow] = [];
          groupOrder.push(rcLow);
        }
        groupMap[rcLow].push(r);
      } else {
        otherRows.push(r);
      }
    });

    var idx = groupOrder.indexOf(cLow);
    if (idx < 0) return getData(ym || null);
    var swapIdx = dir === 'up' ? idx - 1 : idx + 1;
    if (swapIdx < 0 || swapIdx >= groupOrder.length) return getData(ym || null);

    var tmp = groupOrder[idx];
    groupOrder[idx] = groupOrder[swapIdx];
    groupOrder[swapIdx] = tmp;

    var rebuilt = [];
    groupOrder.forEach(function(k) {
      groupMap[k].forEach(function(r) { rebuilt.push(r); });
    });
    otherRows.forEach(function(r) { rebuilt.push(r); });

    if (rebuilt.length) {
      sh.getRange(2, 1, rebuilt.length, 6).setValues(rebuilt);
    }
    invalidateDataCache_();
    return getData(ym || null);
  } finally {
    lock.releaseLock();
  }
}

function saveCategoryOrder(type, orderedTree, ym) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var env = ensureSheets_();
    var sh = env.katSh;
    var last = sh.getLastRow();
    if (last <= 1 || !Array.isArray(orderedTree) || !orderedTree.length) return getData(ym || null);
    var rows = sh.getRange(2, 1, last - 1, 6).getValues();
    var tLow = String(type || 'Pengeluaran').trim().toLowerCase();

    var otherRows = [];
    var catRowsMap = {};
    rows.forEach(function(r) {
      var rt = String(r[0] || '').trim().toLowerCase();
      var rc = String(r[1] || '').trim();
      var rcLow = rc.toLowerCase();
      var ra = String(r[5] || 'YA').trim().toUpperCase();
      if (rt === tLow && rc && ra !== 'TIDAK') {
        if (!catRowsMap[rcLow]) catRowsMap[rcLow] = [];
        catRowsMap[rcLow].push(r);
      } else {
        otherRows.push(r);
      }
    });

    var rebuilt = [];
    var usedCats = {};
    orderedTree.forEach(function(item) {
      var cLow = String((item && item.n) || '').trim().toLowerCase();
      if (!cLow || !catRowsMap[cLow] || usedCats[cLow]) return;
      usedCats[cLow] = true;
      var group = catRowsMap[cLow];
      var subOrder = Array.isArray(item.subs) ? item.subs.map(function(s) { return String(s || '').trim().toLowerCase(); }) : [];
      if (subOrder.length) {
        var subMap = {}, noSubRows = [];
        group.forEach(function(r) {
          var sLow = String(r[3] || '').trim().toLowerCase();
          if (!sLow) { noSubRows.push(r); return; }
          if (!subMap[sLow]) subMap[sLow] = [];
          subMap[sLow].push(r);
        });
        noSubRows.forEach(function(r) { rebuilt.push(r); });
        var usedSubs = {};
        subOrder.forEach(function(sl) {
          if (subMap[sl] && !usedSubs[sl]) {
            usedSubs[sl] = true;
            subMap[sl].forEach(function(r) { rebuilt.push(r); });
          }
        });
        Object.keys(subMap).forEach(function(sl) {
          if (!usedSubs[sl]) subMap[sl].forEach(function(r) { rebuilt.push(r); });
        });
      } else {
        group.forEach(function(r) { rebuilt.push(r); });
      }
    });

    Object.keys(catRowsMap).forEach(function(cLow) {
      if (!usedCats[cLow]) catRowsMap[cLow].forEach(function(r) { rebuilt.push(r); });
    });
    otherRows.forEach(function(r) { rebuilt.push(r); });

    if (rebuilt.length) {
      sh.getRange(2, 1, rebuilt.length, 6).setValues(rebuilt);
    }
    invalidateDataCache_();
    return getData(ym || null);
  } finally {
    lock.releaseLock();
  }
}

function deleteCategory(type, cat, sub, ym) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var env = ensureSheets_();
    var sh = env.katSh;
    type = String(type || 'Pengeluaran').trim().toLowerCase();
    cat = String(cat || '').trim().toLowerCase();
    sub = String(sub || '').trim().toLowerCase();
    var last = sh.getLastRow();
    if (last <= 1) throw new Error('Data kategori kosong');
    var rows = sh.getRange(2, 1, last - 1, 6).getValues();
    var updated = false;
    for (var i = 0; i < rows.length; i++) {
      var rType = String(rows[i][0] || '').trim().toLowerCase();
      var rCat = String(rows[i][1] || '').trim().toLowerCase();
      var rSub = String(rows[i][3] || '').trim().toLowerCase();
      if (rType === type && rCat === cat) {
        if (!sub || rSub === sub) {
          sh.getRange(i + 2, 6).setValue('TIDAK');
          updated = true;
        }
      }
    }
    if (!updated) throw new Error('Kategori tidak ditemukan');
    invalidateDataCache_();
    return getData(ym || null);
  } finally {
    lock.releaseLock();
  }
}

function safeText_(s) {
  s = String(s == null ? '' : s);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function validateTx_(tx) {
  if (!tx) throw new Error('Data transaksi kosong');
  var tp = String(tx.type || 'Pengeluaran').trim();
  if (['Pengeluaran', 'Pemasukan', 'Transfer'].indexOf(tp) < 0) throw new Error('Tipe transaksi tidak valid');
  var amt = Number(tx.amount);
  if (!isFinite(amt) || amt <= 0) throw new Error('Jumlah harus lebih dari 0');
  if (amt > 1e12) throw new Error('Jumlah terlalu besar');
  var date = fmtDate_(tx.date);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('Tanggal tidak valid');
  return { type: tp, amount: amt, date: date };
}

function addTransaction(tx) {
  var v = validateTx_(tx);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var env = ensureSheets_();
    var sh = env.txSh;
    if (tx.eventId || tx.currency || tx.fxAmount) ensureTxEventCols_(sh);
    var col = txCols_(sh);
    var id = 'TX-' + new Date().getTime() + '-' + Math.floor(Math.random() * 900 + 100);
    var vals = {
      id: id, date: v.date, type: v.type,
      cat: safeText_(tx.cat || 'Lainnya'),
      sub: safeText_(tx.sub || ''),
      amount: v.amount,
      wallet: tx.wallet || 'Cash',
      walletTo: tx.walletTo || '',
      note: safeText_(tx.note || ''),
      eventId: String(tx.eventId || '').trim(),
      currency: String(tx.currency || '').trim().toUpperCase(),
      fxAmount: Number(tx.fxAmount) || 0
    };
    var n = Math.max(sh.getLastColumn(), 9);
    Object.keys(col).forEach(function(k) { n = Math.max(n, col[k] + 1); });
    var row = [];
    for (var i = 0; i < n; i++) row.push('');
    Object.keys(vals).forEach(function(k) { row[col[k]] = vals[k]; });
    sh.appendRow(row);
    invalidateDataCache_();
    return getData(v.date.slice(0, 7));
  } finally {
    lock.releaseLock();
  }
}

function addMultipleTransactions(txList) {
  if (!txList || !txList.length) throw new Error('Daftar transaksi kosong');
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var env = ensureSheets_();
    var sh = env.txSh;
    if (txList.some(function(x) { return x && (x.eventId || x.currency || x.fxAmount); })) ensureTxEventCols_(sh);
    var col = txCols_(sh);
    var n = Math.max(sh.getLastColumn(), 9);
    Object.keys(col).forEach(function(k) { n = Math.max(n, col[k] + 1); });

    var baseTs = new Date().getTime();
    var rows = [];
    var targetYm = '';

    for (var idx = 0; idx < txList.length; idx++) {
      var tx = txList[idx];
      var amt = Number(tx.amount) || 0;
      if (amt <= 0) continue;
      var date = fmtDate_(tx.date);
      if (!targetYm) targetYm = date.slice(0, 7);
      var id = 'TX-' + (baseTs + idx) + '-' + Math.floor(Math.random() * 900 + 100);
      var vals = {
        id: id,
        date: date,
        type: tx.type || 'Pengeluaran',
        cat: safeText_(tx.cat || 'Lainnya'),
        sub: safeText_(tx.sub || ''),
        amount: amt,
        wallet: tx.wallet || 'Cash',
        walletTo: tx.walletTo || '',
        note: safeText_(tx.note || ''),
        eventId: String(tx.eventId || '').trim(),
        currency: String(tx.currency || '').trim().toUpperCase(),
        fxAmount: Number(tx.fxAmount) || 0
      };
      var row = [];
      for (var c = 0; c < n; c++) row.push('');
      Object.keys(vals).forEach(function(k) { row[col[k]] = vals[k]; });
      rows.push(row);
    }

    if (rows.length) {
      var startRow = sh.getLastRow() + 1;
      sh.getRange(startRow, 1, rows.length, n).setValues(rows);
      invalidateDataCache_();
    }
    return getData(targetYm || null);
  } finally {
    lock.releaseLock();
  }
}

function updateTransaction(tx, currentYm) {
  var v = validateTx_(tx);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var env = ensureSheets_();
    var sh = env.txSh;
    if (tx.eventId || tx.currency || tx.fxAmount) ensureTxEventCols_(sh);
    var col = txCols_(sh);
    var row = findTxRow_(sh, col, tx.id);
    if (row < 2) throw new Error('Transaksi tidak ditemukan. Muat ulang data lalu coba lagi.');
    var n = Math.max(sh.getLastColumn(), 9);
    Object.keys(col).forEach(function(k) { n = Math.max(n, col[k] + 1); });
    var rowVals = sh.getRange(row, 1, 1, n).getValues()[0];
    var vals = {
      date: v.date,
      type: v.type,
      cat: safeText_(tx.cat || 'Lainnya'),
      sub: safeText_(tx.sub || ''),
      amount: v.amount,
      wallet: tx.wallet || 'Cash',
      walletTo: tx.walletTo || '',
      note: safeText_(tx.note || ''),
      eventId: String(tx.eventId || '').trim(),
      currency: String(tx.currency || '').trim().toUpperCase(),
      fxAmount: Number(tx.fxAmount) || 0
    };
    Object.keys(vals).forEach(function(k) { rowVals[col[k]] = vals[k]; });
    sh.getRange(row, 1, 1, n).setValues([rowVals]);
    invalidateDataCache_();
    return getData(currentYm || v.date.slice(0, 7));
  } finally {
    lock.releaseLock();
  }
}

function deleteTransaction(id, currentYm) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var env = ensureSheets_();
    var sh = env.txSh;
    var col = txCols_(sh);
    var row = findTxRow_(sh, col, id);
    if (row < 2) throw new Error('Transaksi tidak ditemukan (mungkin sudah terhapus). Muat ulang data.');
    sh.deleteRow(row);
    invalidateDataCache_();
    return getData(currentYm || null);
  } finally {
    lock.releaseLock();
  }
}

function setBudget(cat, amount, ym) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var env = ensureSheets_();
    var sh = env.budSh;
    var katCfg = readKategoriCfg_(env.katSh);
    var targetYm = ym || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
    var amt = Math.max(0, Math.round(Number(amount) || 0));
    var last = sh.getLastRow();
    if (last > 1) {
      var rows = sh.getRange(2, 1, last - 1, 3).getValues();
      for (var i = rows.length - 1; i >= 0; i--) {
        if (normYm_(rows[i][0]) === targetYm && canonCat_('Pengeluaran', String(rows[i][1]).trim(), katCfg) === String(cat).trim()) {
          sh.getRange(i + 2, 3).setValue(amt);
          invalidateDataCache_();
          return getData(targetYm);
        }
      }
    }
    sh.appendRow([targetYm, cat, amt]);
    sh.getRange(sh.getLastRow(), 1).setNumberFormat('@').setValue(targetYm);
    invalidateDataCache_();
    return getData(targetYm);
  } finally {
    lock.releaseLock();
  }
}

function setTarget(cat, amount, ym) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var allowed = ['Dana Darurat', 'Reksadana', 'Saham'];
    cat = String(cat || '').trim();
    if (allowed.indexOf(cat) < 0) throw new Error('Aset tidak dikenal');
    var env = ensureSheets_();
    var sh = env.tgSh;
    var amt = Math.max(0, Math.round(Number(amount) || 0));
    var last = sh.getLastRow();
    if (last > 1) {
      var rows = sh.getRange(2, 1, last - 1, 1).getValues();
      for (var i = rows.length - 1; i >= 0; i--) {
        if (String(rows[i][0]).trim() === cat) {
          sh.getRange(i + 2, 2).setValue(amt);
          invalidateDataCache_();
          return getData(ym || null);
        }
      }
    }
    sh.appendRow([cat, amt]);
    invalidateDataCache_();
    return getData(ym || null);
  } finally {
    lock.releaseLock();
  }
}

function setOpening(walletName, amount, ym) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var env = ensureSheets_();
    var sh = env.saSh;
    var targetYm = ym || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
    var amt = Math.round(Number(amount) || 0);
    var last = sh.getLastRow();
    if (last > 1) {
      var rows = sh.getRange(2, 1, last - 1, 3).getValues();
      for (var i = 0; i < rows.length; i++) {
        if (normYm_(rows[i][0]) === targetYm && String(rows[i][1]).trim() === String(walletName).trim()) {
          sh.getRange(i + 2, 3).setValue(amt);
          invalidateDataCache_();
          return getData(targetYm);
        }
      }
    }
    sh.appendRow([targetYm, walletName, amt]);
    sh.getRange(sh.getLastRow(), 1).setNumberFormat('@').setValue(targetYm);
    invalidateDataCache_();
    return getData(targetYm);
  } finally {
    lock.releaseLock();
  }
}

function exportAllData() {
  var env = ensureSheets_();
  var sheets = [
    { name: 'Transaksi', sh: env.txSh },
    { name: 'Budget', sh: env.budSh },
    { name: 'SaldoAwal', sh: env.saSh },
    { name: 'Dompet', sh: env.dompetSh },
    { name: 'Kategori', sh: env.katSh },
    { name: 'Events', sh: env.evSh }
  ];
  var out = {};
  sheets.forEach(function(item) {
    var sh = item.sh;
    var r = sh.getLastRow();
    var c = sh.getLastColumn();
    if (r < 1 || c < 1) {
      out[item.name] = [];
      return;
    }
    var vals = sh.getRange(1, 1, r, c).getValues();
    out[item.name] = vals.map(function(row) {
      return row.map(function(cell) {
        if (Object.prototype.toString.call(cell) === '[object Date]' && !isNaN(cell.getTime())) {
          return Utilities.formatDate(cell, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        }
        return cell === null || cell === undefined ? '' : cell;
      });
    });
  });
  return out;
}

/* ===== Fungsi bantu (dijalankan manual dari editor Apps Script) ===== */
function tes() {
  var d = getData('2026-11');
  Logger.log('Pemasukan: ' + d.income);
  Logger.log('Pengeluaran: ' + d.expense);
  Logger.log('Budget: ' + JSON.stringify(d.budgets));
  Logger.log('Budget Alerts: ' + JSON.stringify(d.budgetAlerts));
  Logger.log('Saldo dompet: ' + JSON.stringify(d.wallets.map(function(w){ return w.name + '=' + w.balance; })));
}

function rapikanKategori() {
  var env = ensureSheets_();
  var sh = env.txSh, col = txCols_(sh), last = sh.getLastRow();
  var katCfg = readKategoriCfg_(env.katSh);
  if (last <= 1) return;
  var n = Math.max(sh.getLastColumn(), 9);
  var v = sh.getRange(2, 1, last - 1, n).getValues();
  var cats = [], subs = [];
  v.forEach(function(r) {
    var tp = String(r[col.type] || 'Pengeluaran').trim();
    var c = canonCat_(tp, String(r[col.cat] || '').trim(), katCfg);
    cats.push([c]);
    subs.push([canonSub_(tp, c, String(r[col.sub] || '').trim(), katCfg)]);
  });
  sh.getRange(2, col.cat + 1, last - 1, 1).setValues(cats);
  sh.getRange(2, col.sub + 1, last - 1, 1).setValues(subs);
  invalidateDataCache_();
}

function beriIdBarisLama() {
  var env = ensureSheets_();
  var sh = env.txSh, col = txCols_(sh), last = sh.getLastRow();
  if (last <= 1) return;
  var idRng = sh.getRange(2, col.id + 1, last - 1, 1);
  var ids = idRng.getValues();
  var dates = sh.getRange(2, col.date + 1, last - 1, 1).getValues();
  var base = new Date().getTime(), n = 0;
  for (var i = 0; i < ids.length; i++) {
    if (dates[i][0] && !String(ids[i][0] || '').trim()) {
      ids[i][0] = 'TX-' + (base + i) + '-' + Math.floor(Math.random() * 900 + 100);
      n++;
    }
  }
  if (n) {
    idRng.setValues(ids);
    invalidateDataCache_();
  }
  Logger.log('ID baru dibuat untuk ' + n + ' baris');
}

function migrasiKategoriBaru() {
  var env = ensureSheets_();
  var sh = env.txSh, col = txCols_(sh), last = sh.getLastRow();
  var katCfg = readKategoriCfg_(env.katSh);
  if (last <= 1) return;
  var n = Math.max(sh.getLastColumn(), 9);
  var v = sh.getRange(2, 1, last - 1, n).getValues();
  var cats = [], subs = [];
  var mapMakan = {
    'makanpagi': 'Makanan', 'makansiang': 'Makanan', 'makanmalam': 'Makanan',
    'jajankopi': 'Cafe'
  };
  v.forEach(function(r) {
    var tp = String(r[col.type] || 'Pengeluaran').trim();
    var c = canonCat_(tp, String(r[col.cat] || '').trim(), katCfg);
    var s = String(r[col.sub] || '').trim();
    if (tp === 'Pengeluaran') {
      var k = slug_(s);
      if (c === 'Tagihan' && (k === 'aisubs' || k === 'appseditor')) c = 'Produktivitas';
      if (c === 'Makan' && mapMakan[k]) s = mapMakan[k];
    }
    cats.push([c]);
    subs.push([s]);
  });
  sh.getRange(2, col.cat + 1, last - 1, 1).setValues(cats);
  sh.getRange(2, col.sub + 1, last - 1, 1).setValues(subs);
  invalidateDataCache_();
}

function scanQuotaKey_() {
  return 'SCAN_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
}
function scanUsedToday_() {
  return Number(PropertiesService.getScriptProperties().getProperty(scanQuotaKey_())) || 0;
}
function scanCountUp_() {
  var props = PropertiesService.getScriptProperties();
  var key = scanQuotaKey_();
  var all = props.getProperties();
  Object.keys(all).forEach(function(k) { // hapus hitungan hari-hari lama
    if (k.indexOf('SCAN_') === 0 && k !== key) props.deleteProperty(k);
  });
  props.setProperty(key, String((Number(all[key]) || 0) + 1));
}
function tesSisaScan() {
  Logger.log('Terpakai hari ini: ' + scanUsedToday_() + ' dari ' + SCAN_LIMIT_PER_DAY);
}

/* ===== SMART SCAN STRUK & BUKTI TRANSFER (Gemini + Auto-Split Multi-Kategori) ===== */
function scanReceipt(base64Data, mimeType) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('ERR_NO_KEY');
  apiKey = apiKey.trim();
  if (scanUsedToday_() >= SCAN_LIMIT_PER_DAY) throw new Error('ERR_SCAN_LIMIT');

  var env = ensureSheets_();
  var dompetList = readDompetList_(env.dompetSh).map(function(d) { return d.n; });
  var katCfg = readKategoriCfg_(env.katSh);
  var tz = Session.getScriptTimeZone();
  var today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  var katListStr = ['keluar', 'masuk', 'transfer'].map(function(k) {
    var tp = k === 'masuk' ? 'Pemasukan' : (k === 'transfer' ? 'Transfer' : 'Pengeluaran');
    var items = (katCfg[k] || []).map(function(c) {
      var subs = (c.sub || []).map(function(s) { return s.n; }).join(', ');
      return c.n + (subs ? ' [' + subs + ']' : '');
    }).join('; ');
    return tp + ': ' + items;
  }).join('\n');

  var prompt =
    'Kamu adalah asisten pencatat keuangan pribadi. Ekstrak data transaksi dari gambar struk belanja, nota, QRIS, atau bukti transfer m-banking ini.\n' +
    'Tanggal hari ini: ' + today + '\n' +
    'Daftar Dompet tersedia: ' + dompetList.join(', ') + '\n' +
    'Daftar Kategori & [SubKategori] tersedia:\n' + katListStr + '\n\n' +
    'Aturan wajib:\n' +
    '1. "type": pilih salah satu dari "Pengeluaran", "Pemasukan", atau "Transfer".\n' +
    '2. "totalAmount": angka total pembayaran bersih akhir di struk (integer tanpa titik/koma, contoh: 37000).\n' +
    '3. "wallet": pilih nama Dompet yang paling cocok dari daftar dompet (misal BCA, Gopay, Jago, atau "Cash" jika tunai/tidak jelas).\n' +
    '4. "walletTo": isi nama Dompet tujuan HANYA jika type="Transfer" dan cat="Antar Dompet", selain itu "".\n' +
    '5. "date": format YYYY-MM-DD dari tanggal di struk; jika tidak terbaca/tidak ada tahun, gunakan "' + today + '".\n' +
    '6. "items": array berisi kelompok transaksi berdasarkan kategori.\n' +
    '   - Jika semua barang di struk masuk ke 1 kategori/subkategori yang sama (atau ini bukti transfer), cukup buat 1 elemen di dalam "items".\n' +
    '   - Jika dalam 1 struk ada barang dari kategori berbeda (misal snack/minuman dan obat/vitamin, atau makanan dan kebutuhan rumah), PECAH dan gabungkan per (cat + sub) yang sama menjadi beberapa elemen di "items".\n' +
    '   - Distribusikan diskon/voucher/PPN/biaya layanan secara proporsional sehingga jumlah seluruh "amount" di "items" sama persis dengan "totalAmount".\n' +
    '   - Setiap objek di "items" wajib punya:\n' +
    '     * "cat": nama Kategori persis dari daftar di atas.\n' +
    '     * "sub": nama SubKategori persis dari dalam tanda [] kategori tersebut, atau "" jika tidak ada.\n' +
    '     * "amount": integer nominal kelompok kategori tersebut.\n' +
    '     * "note": nama merchant/toko singkat + nama item di kelompok ini (maks 50 karakter).\n' +
    'Balas HANYA dengan JSON valid tanpa markdown.';

  var models = ['gemini-3.1-flash-lite', 'gemini-3.5-flash-lite', 'gemini-3.5-flash'];

  var res, code = 0, bodyText = '';

  for (var i = 0; i < models.length; i++) {
    var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + models[i] + ':generateContent';
    var payload = {
      contents: [{
        parts: [
          { text: prompt },
          { inline_data: { mime_type: mimeType || 'image/jpeg', data: base64Data } }
        ]
      }],
      generationConfig: { temperature: 0.1, responseMimeType: 'application/json' }
    };

    res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-goog-api-key': apiKey },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    code = res.getResponseCode();
    bodyText = res.getContentText();
    if (code === 200) break;
    if (code === 401 || code === 403) break;
  }

  if (code !== 200) {
    var errDetail = String(code);
    try {
      var errObj = JSON.parse(bodyText);
      if (errObj.error && errObj.error.message) {
        errDetail = code + ' - ' + errObj.error.message.slice(0, 70);
      }
    } catch (ignore) {}
    throw new Error('ERR_AI_FAIL:' + errDetail);
  }
  scanCountUp_();

  var json = JSON.parse(bodyText);
  var parts = json.candidates && json.candidates[0] && json.candidates[0].content && json.candidates[0].content.parts;
  var text = parts ? parts.map(function(p) { return p.text || ''; }).join('') : '';
  if (!text) throw new Error('ERR_AI_EMPTY');

  var parsed;
  try {
    var cleanJson = text.replace(/```json|```/g, '').trim();
    var firstBrace = cleanJson.indexOf('{');
    var lastBrace = cleanJson.lastIndexOf('}');
    if (firstBrace >= 0 && lastBrace > firstBrace) {
      cleanJson = cleanJson.slice(firstBrace, lastBrace + 1);
    }
    parsed = JSON.parse(cleanJson);
  } catch (e) {
    throw new Error('ERR_AI_PARSE');
  }

  var tp = ['Pengeluaran', 'Pemasukan', 'Transfer'].indexOf(parsed.type) >= 0 ? parsed.type : 'Pengeluaran';
  var w = String(parsed.wallet || '').trim();
  if (dompetList.indexOf(w) < 0) w = dompetList[0] || 'Cash';
  var wTo = String(parsed.walletTo || '').trim();
  if (dompetList.indexOf(wTo) < 0) wTo = '';
  var dateStr = /^\d{4}-\d{2}-\d{2}$/.test(String(parsed.date || '')) ? parsed.date : today;
  var totalReceipt = Math.max(0, Math.round(Number(parsed.totalAmount || parsed.amount) || 0));

  var rawItems = Array.isArray(parsed.items) && parsed.items.length
    ? parsed.items
    : [{ cat: parsed.cat, sub: parsed.sub, amount: totalReceipt, note: parsed.note }];

  var merged = [];
  var mergeMap = {};
  rawItems.forEach(function(it) {
    var c = canonCat_(tp, String(it.cat || '').trim(), katCfg) || 'Lainnya';
    var s = canonSub_(tp, c, String(it.sub || '').trim(), katCfg) || '';
    var a = Math.max(0, Math.round(Number(it.amount) || 0));
    var nt = String(it.note || '').trim().slice(0, 100);
    if (a <= 0 && rawItems.length > 1) return;
    var key = c + '|' + s;
    if (mergeMap[key] !== undefined) {
      var idx = mergeMap[key];
      merged[idx].amount += a;
      if (nt && merged[idx].note.indexOf(nt) < 0) {
        merged[idx].note = (merged[idx].note + ', ' + nt).slice(0, 100);
      }
    } else {
      mergeMap[key] = merged.length;
      merged.push({ cat: c, sub: s, amount: a, note: nt });
    }
  });

  if (!merged.length) {
    var defCat = canonCat_(tp, String(parsed.cat || '').trim(), katCfg) || 'Lainnya';
    merged.push({
      cat: defCat,
      sub: canonSub_(tp, defCat, String(parsed.sub || '').trim(), katCfg) || '',
      amount: totalReceipt,
      note: String(parsed.note || '').trim().slice(0, 100)
    });
  }

  if (totalReceipt > 0 && merged.length > 1) {
    var sumItems = merged.reduce(function(acc, x) { return acc + x.amount; }, 0);
    var diff = totalReceipt - sumItems;
    if (diff !== 0 && Math.abs(diff) <= Math.max(500, totalReceipt * 0.15)) {
      var maxIdx = 0;
      for (var m = 1; m < merged.length; m++) {
        if (merged[m].amount > merged[maxIdx].amount) maxIdx = m;
      }
      if (merged[maxIdx].amount + diff > 0) {
        merged[maxIdx].amount += diff;
      }
    }
  }

  var first = merged[0];
  return {
    type: tp,
    amount: merged.length === 1 ? first.amount : totalReceipt,
    cat: first.cat,
    sub: first.sub,
    wallet: w,
    walletTo: wTo,
    date: dateStr,
    note: first.note,
    items: merged
  };
}

/* ===== AI SUARA (REKAM AUDIO -> GEMINI) ===== */
function parseVoiceTransaction(base64Audio, mimeType, textInput) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('ERR_NO_KEY');
  apiKey = apiKey.trim();

  var env = ensureSheets_();
  var dompetList = readDompetList_(env.dompetSh).map(function(d) { return d.n; });
  var katCfg = readKategoriCfg_(env.katSh);
  var tz = Session.getScriptTimeZone();
  var today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  var katListStr = ['keluar', 'masuk', 'transfer'].map(function(k) {
    var tp = k === 'masuk' ? 'Pemasukan' : (k === 'transfer' ? 'Transfer' : 'Pengeluaran');
    var items = (katCfg[k] || []).map(function(c) {
      var subs = (c.sub || []).map(function(s) { return s.n; }).join(', ');
      return c.n + (subs ? ' [' + subs + ']' : '');
    }).join('; ');
    return tp + ': ' + items;
  }).join('\n');

  var prompt =
    'Kamu adalah asisten pencatat keuangan pribadi. Dengarkan rekaman suara transaksi berbahasa Indonesia/Inggris ini dan ekstrak menjadi 1 transaksi.\n' +
    'Tanggal hari ini: ' + today + '\n' +
    'Daftar Dompet tersedia: ' + dompetList.join(', ') + '\n' +
    'Daftar Kategori & [SubKategori] tersedia:\n' + katListStr + '\n\n' +
    'Aturan wajib:\n' +
    '1. "type": pilih salah satu dari "Pengeluaran", "Pemasukan", atau "Transfer" (default "Pengeluaran").\n' +
    '2. "amount": angka nominal dalam Rupiah sebagai integer tanpa titik/koma (contoh: "5 ribu" -> 5000, "dua puluh lima ribu" -> 25000).\n' +
    '3. "cat": pilih nama Kategori yang paling cocok persis dari daftar di atas.\n' +
    '4. "sub": pilih nama SubKategori yang paling cocok persis dari dalam tanda [] kategori tersebut, atau "" jika tidak ada.\n' +
    '5. "wallet": pilih nama Dompet yang disebut atau paling cocok dari daftar dompet (misal "warung kaki lima" -> "Cash" jika ada di daftar).\n' +
    '6. "walletTo": isi nama Dompet tujuan HANYA jika type="Transfer" dan cat="Antar Dompet", selain itu "".\n' +
    '7. "date": format YYYY-MM-DD jika disebut (misal "kemarin"), jika tidak disebut gunakan "' + today + '".\n' +
    '8. "note": ringkasan singkat barang/kegiatan + tempat (maks 60 karakter).\n' +
    'Balas HANYA dengan JSON valid tanpa markdown.';

  var models = ['gemini-3.1-flash-lite', 'gemini-3.5-flash-lite', 'gemini-3.5-flash'];
  var cleanAudio = String(base64Audio || '').replace(/^data:[^;]+;base64,/, '');
  var cleanMime = String(mimeType || 'audio/webm').split(';')[0].trim() || 'audio/webm';
  var res, code = 0, bodyText = '';

  var teks = String(textInput || '').trim().slice(0, 500);
  var inputParts = [{ text: prompt }];
  if (teks) {
    inputParts.push({ text: 'Ucapan pengguna: ' + teks });
  } else {
    inputParts.push({ inline_data: { mime_type: cleanMime, data: cleanAudio } });
  }

  for (var i = 0; i < models.length; i++) {
    var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + models[i] + ':generateContent';
    var payload = {
      contents: [{ parts: inputParts }],
      generationConfig: { temperature: 0.1, responseMimeType: 'application/json' }
    };

    res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-goog-api-key': apiKey },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    code = res.getResponseCode();
    bodyText = res.getContentText();
    if (code === 200) break;
    if (code === 401 || code === 403) break;
  }

  if (code !== 200) {
    var errDetail = String(code);
    try {
      var errObj = JSON.parse(bodyText);
      if (errObj.error && errObj.error.message) {
        errDetail = code + ' - ' + errObj.error.message.slice(0, 70);
      }
    } catch (ignore) {}
    throw new Error('ERR_AI_FAIL:' + errDetail);
  }

  scanCountUp_();
  var json = JSON.parse(bodyText);
  var parts = json.candidates && json.candidates[0] && json.candidates[0].content && json.candidates[0].content.parts;
  var text = parts ? parts.map(function(p) { return p.text || ''; }).join('') : '';
  if (!text) throw new Error('ERR_AI_EMPTY');

  var parsed;
  try {
    var cleanJson = text.replace(/```json|```/g, '').trim();
    var firstBrace = cleanJson.indexOf('{');
    var lastBrace = cleanJson.lastIndexOf('}');
    if (firstBrace >= 0 && lastBrace > firstBrace) cleanJson = cleanJson.slice(firstBrace, lastBrace + 1);
    parsed = JSON.parse(cleanJson);
  } catch (e) {
    throw new Error('ERR_AI_PARSE');
  }

  var tp = ['Pengeluaran', 'Pemasukan', 'Transfer'].indexOf(parsed.type) >= 0 ? parsed.type : 'Pengeluaran';
  var cat = canonCat_(tp, String(parsed.cat || '').trim(), katCfg);
  var sub = canonSub_(tp, cat, String(parsed.sub || '').trim(), katCfg);
  var w = String(parsed.wallet || '').trim();
  if (dompetList.indexOf(w) < 0) w = dompetList[0] || 'Cash';
  var wTo = String(parsed.walletTo || '').trim();
  if (dompetList.indexOf(wTo) < 0) wTo = '';

  return {
    type: tp,
    amount: Math.max(0, Math.round(Number(parsed.amount) || 0)),
    cat: cat,
    sub: sub || '',
    wallet: w,
    walletTo: wTo,
    date: /^\d{4}-\d{2}-\d{2}$/.test(String(parsed.date || '')) ? parsed.date : today,
    note: String(parsed.note || '').trim().slice(0, 100)
  };
}

function tesGemini() {
  var key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) { Logger.log('GEMINI_API_KEY belum diisi di Script Properties'); return; }

  var res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models?pageSize=100', {
    method: 'get',
    headers: { 'x-goog-api-key': key },
    muteHttpExceptions: true
  });

  var code = res.getResponseCode();
  Logger.log('Status: ' + code);
  if (code !== 200) { Logger.log(res.getContentText().slice(0, 500)); return; }

  var models = JSON.parse(res.getContentText()).models || [];
  models.filter(function(m) {
    return /flash/i.test(m.name) &&
           (m.supportedGenerationMethods || []).indexOf('generateContent') >= 0;
  }).forEach(function(m) { Logger.log(m.name); });
}

function tesModelGemini() {
  var key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) { Logger.log('GEMINI_API_KEY belum diisi'); return; }

  var models = [
    'gemini-2.5-flash', 'gemini-2.5-flash-lite',
    'gemini-3.1-flash-lite', 'gemini-3.5-flash', 'gemini-3.5-flash-lite',
    'gemini-3.6-flash', 'gemini-3.7-flash', 'gemini-3.8-flash',
    'gemini-flash-latest', 'gemini-flash-lite-latest'
  ];

  models.forEach(function(m) {
    var t0 = new Date().getTime();
    var res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + m + ':generateContent', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-goog-api-key': key },
      payload: JSON.stringify({ contents: [{ parts: [{ text: 'Balas satu kata: OK' }] }] }),
      muteHttpExceptions: true
    });
    var code = res.getResponseCode();
    var msg = '';
    if (code !== 200) {
      try { msg = JSON.parse(res.getContentText()).error.message.slice(0, 80); }
      catch (e) { msg = res.getContentText().slice(0, 80); }
    }
    Logger.log(m + ' → ' + code + ' (' + (new Date().getTime() - t0) + ' ms) ' + msg);
  });
}

/* ===== MANAJEMEN EVENT & TRIP ===== */
function saveEvent(payload, ym) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var env = ensureSheets_();
    var sh = env.evSh;
    var name = String((payload && payload.name) || '').trim();
    if (!name) throw new Error('Nama Event tidak boleh kosong');
    var icon = String((payload && payload.icon) || '✈️').trim() || '✈️';
    var currency = String((payload && payload.currency) || 'IDR').trim().toUpperCase() || 'IDR';
    var rate = currency === 'IDR' ? 1 : (Number(payload && payload.rate) || 1);
    var budget = Math.max(0, Math.round(Number(payload && payload.budget) || 0));
    var startDate = payload && payload.startDate ? fmtDate_(payload.startDate) : '';
    var endDate = payload && payload.endDate ? fmtDate_(payload.endDate) : '';
    var excl = String((payload && payload.excludeReport) || 'TIDAK').trim().toUpperCase() === 'YA' ? 'YA' : 'TIDAK';
    var id = String((payload && payload.id) || '').trim();

    if (id) {
      var last = sh.getLastRow();
      if (last > 1) {
        var rows = sh.getRange(2, 1, last - 1, 10).getValues();
        for (var i = 0; i < rows.length; i++) {
          if (String(rows[i][0]).trim() === id) {
            var curStatus = String(rows[i][8] || 'AKTIF').trim().toUpperCase();
            sh.getRange(i + 2, 1, 1, 10).setValues([[id, safeText_(name), icon, currency, rate, budget, startDate, endDate, curStatus, excl]]);
            invalidateDataCache_();
            return getData(ym || null);
          }
        }
      }
    }
    var newId = 'EV-' + new Date().getTime() + '-' + Math.floor(Math.random() * 900 + 100);
    sh.appendRow([newId, safeText_(name), icon, currency, rate, budget, startDate, endDate, 'AKTIF', excl]);
    invalidateDataCache_();
    return getData(ym || null);
  } finally {
    lock.releaseLock();
  }
}

function toggleEventStatus(eventId, status, ym) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var env = ensureSheets_();
    var sh = env.evSh;
    var last = sh.getLastRow();
    var next = String(status || '').trim().toUpperCase() === 'SELESAI' ? 'SELESAI' : 'AKTIF';
    if (last > 1) {
      var ids = sh.getRange(2, 1, last - 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) {
        if (String(ids[i][0]).trim() === String(eventId).trim()) {
          sh.getRange(i + 2, 9).setValue(next);
          invalidateDataCache_();
          return getData(ym || null);
        }
      }
    }
    throw new Error('Event tidak ditemukan');
  } finally {
    lock.releaseLock();
  }
}

function deleteEvent(eventId, ym) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var env = ensureSheets_();
    var sh = env.evSh;
    var last = sh.getLastRow();
    if (last > 1) {
      var ids = sh.getRange(2, 1, last - 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) {
        if (String(ids[i][0]).trim() === String(eventId).trim()) {
          sh.deleteRow(i + 2);
          invalidateDataCache_();
          return getData(ym || null);
        }
      }
    }
    throw new Error('Event tidak ditemukan');
  } finally {
    lock.releaseLock();
  }
}

function getEventTransactions(eventId) {
  var env = ensureSheets_();
  var katCfg = readKategoriCfg_(env.katSh);
  var tCol = txCols_(env.txSh);
  var txLastRow = env.txSh.getLastRow();
  if (txLastRow <= 1) return [];
  var txColsCount = Math.max(12, env.txSh.getLastColumn());
  var tRows = env.txSh.getRange(2, 1, txLastRow - 1, txColsCount).getValues();
  var target = String(eventId || '').trim();
  var out = [];
  for (var i = 0; i < tRows.length; i++) {
    var r = tRows[i];
    if (String(r[tCol.eventId] || '').trim() !== target) continue;
    var ds = fmtDate_(r[tCol.date]);
    if (!ds || ds.length < 7) continue;
    var tp = String(r[tCol.type] || 'Pengeluaran').trim();
    var cat = canonCat_(tp, String(r[tCol.cat] || '').trim(), katCfg);
    var sub = canonSub_(tp, cat, String(r[tCol.sub] || '').trim(), katCfg);
    out.push({
      id: String(r[tCol.id] || ('ROW-' + (i + 2))).trim(),
      date: ds,
      type: tp,
      cat: cat,
      sub: sub || null,
      amount: parseAmount_(r[tCol.amount]),
      wallet: String(r[tCol.wallet] || 'Cash').trim() || 'Cash',
      walletTo: String(r[tCol.walletTo] || '').trim() || null,
      note: String(r[tCol.note] || '').trim(),
      eventId: target,
      currency: String(r[tCol.currency] || '').trim().toUpperCase(),
      fxAmount: parseAmount_(r[tCol.fxAmount])
    });
  }
  out.sort(function(a, b) {
    if (a.date !== b.date) return a.date < b.date ? 1 : -1;
    return a.id < b.id ? 1 : -1;
  });
  return out;
}

/* ===== KURS VALAS OTOMATIS (LIVE EXCHANGE RATE) ===== */
function getLiveExchangeRate(currency) {
  var code = String(currency || 'IDR').trim().toUpperCase().replace(/[^A-Z]/g, '');
  if (!code || code === 'IDR') return { currency: 'IDR', rate: 1, updated: '' };
  var isoMap = { RMB: 'CNY', NTD: 'TWD' };
  var iso = isoMap[code] || code;
  var cache = CacheService.getScriptCache();
  var cacheKey = 'FX_RATE_' + iso;
  var cached = cache.get(cacheKey);
  if (cached) {
    try {
      var cObj = JSON.parse(cached);
      cObj.currency = code;
      return cObj;
    } catch (ignore) {}
  }
  var res = UrlFetchApp.fetch('https://open.er-api.com/v6/latest/' + encodeURIComponent(iso), {
    method: 'get',
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) throw new Error('Gagal mengambil kurs live untuk ' + code);
  var json = JSON.parse(res.getContentText());
  var idrRate = json && json.rates && Number(json.rates.IDR);
  if (!isFinite(idrRate) || idrRate <= 0) throw new Error('Kurs ' + code + ' ke IDR tidak tersedia');
  var rounded = idrRate < 100 ? Math.round(idrRate * 100) / 100 : Math.round(idrRate);
  var out = {
    currency: code,
    rate: rounded,
    updated: String(json.time_last_update_utc || '').slice(0, 16)
  };
  try { cache.put(cacheKey, JSON.stringify(out), 21600); } catch (ignore) {}
  return out;
}

/* ===== AI FINANCIAL ADVISOR / REVIEW BULANAN (GEMINI) ===== */
function analyzeMonthWithAI(ym, lang) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('ERR_NO_KEY');
  apiKey = apiKey.trim();

  var d = getData(ym);
  var isEn = String(lang || 'id').toLowerCase() === 'en';
  var topCats = (d.categories || []).slice(0, 6).map(function(c) {
    return c.cat + ': Rp ' + Math.round(c.amount).toLocaleString('id-ID');
  }).join(', ') || '-';
  var alerts = (d.budgetAlerts || []).map(function(a) {
    return a.cat + ' (' + a.pct + '% dari budget Rp ' + Math.round(a.limit).toLocaleString('id-ID') + ')';
  }).join(', ') || 'Tidak ada';
  var activeEvs = (d.events || []).filter(function(e) { return e.status !== 'SELESAI'; }).map(function(e) {
    return e.name + ' (Rp ' + Math.round(e.spent || 0).toLocaleString('id-ID') + ')';
  }).join(', ') || 'Tidak ada';

  var summaryText =
    'Bulan: ' + d.ym + '\n' +
    'Pemasukan: Rp ' + Math.round(d.income || 0).toLocaleString('id-ID') + '\n' +
    'Pengeluaran: Rp ' + Math.round(d.expense || 0).toLocaleString('id-ID') + '\n' +
    'Surplus/Defisit: Rp ' + Math.round(d.saving || 0).toLocaleString('id-ID') + '\n' +
    'Saving Rate: ' + Math.round((d.rate || 0) * 100) + '%\n' +
    'Rata-rata pengeluaran 3 bulan terakhir (pada tanggal yg sama): Rp ' + Math.round((d.report && d.report.avgTotal) || 0).toLocaleString('id-ID') + '\n' +
    'Kas Bebas: Rp ' + Math.round(d.freeCash || 0).toLocaleString('id-ID') + ' | Total Aset: Rp ' + Math.round(d.totalAsset || 0).toLocaleString('id-ID') + '\n' +
    'Pengeluaran Teratas: ' + topCats + '\n' +
    'Peringatan Budget: ' + alerts + '\n' +
    'Event Aktif: ' + activeEvs;

  var prompt = isEn
    ? ('You are a friendly, sharp personal finance advisor. Analyze this monthly financial summary and respond ONLY in valid JSON without markdown:\n' +
       '{"score": 1-100, "headline": "short 1-sentence verdict", "highlights": ["2-3 concise positive points"], "warnings": ["1-3 concise overspending/risk points"], "tips": ["2-3 concrete actionable tips for next week/month"]}\n\nData:\n' + summaryText)
    : ('Kamu adalah penasihat keuangan pribadi yang ramah, jujur, dan tajam. Analisa ringkasan keuangan bulanan berikut dan balas HANYA dengan JSON valid tanpa markdown:\n' +
       '{"score": angka 1-100, "headline": "1 kalimat kesimpulan singkat kondisi bulan ini", "highlights": ["2-3 poin positif singkat"], "warnings": ["1-3 poin waspada/boros singkat"], "tips": ["2-3 saran konkret & praktis untuk minggu/bulan depan"]}\n\nData:\n' + summaryText);

  var models = ['gemini-3.1-flash-lite', 'gemini-3.5-flash-lite', 'gemini-3.5-flash', 'gemini-2.5-flash'];
  var res, code = 0, bodyText = '';
  for (var i = 0; i < models.length; i++) {
    var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + models[i] + ':generateContent';
    res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-goog-api-key': apiKey },
      payload: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { temperature: 0.3, responseMimeType: 'application/json' }
      }),
      muteHttpExceptions: true
    });
    code = res.getResponseCode();
    bodyText = res.getContentText();
    if (code === 200 || code === 401 || code === 403) break;
  }
  if (code !== 200) throw new Error('ERR_AI_FAIL:' + code);

  var json = JSON.parse(bodyText);
  var parts = json.candidates && json.candidates[0] && json.candidates[0].content && json.candidates[0].content.parts;
  var text = parts ? parts.map(function(p) { return p.text || ''; }).join('') : '';
  if (!text) throw new Error('ERR_AI_EMPTY');
  var clean = text.replace(/```json|```/g, '').trim();
  var fb = clean.indexOf('{'), lb = clean.lastIndexOf('}');
  if (fb >= 0 && lb > fb) clean = clean.slice(fb, lb + 1);
  var parsed = JSON.parse(clean);
  return {
    ym: d.ym,
    score: Math.max(0, Math.min(100, Math.round(Number(parsed.score) || 75))),
    headline: String(parsed.headline || '').trim(),
    highlights: Array.isArray(parsed.highlights) ? parsed.highlights.slice(0, 4) : [],
    warnings: Array.isArray(parsed.warnings) ? parsed.warnings.slice(0, 4) : [],
    tips: Array.isArray(parsed.tips) ? parsed.tips.slice(0, 4) : []
  };
}

/* ===== API UNTUK HALAMAN DI GITHUB PAGES ===== */
var API_MAP_ = {
  getData: getData,
  searchTransactions: searchTransactions,
  addWallet: addWallet,
  deleteWallet: deleteWallet,
  addCategory: addCategory,
  reorderCategory: reorderCategory,
  saveCategoryOrder: saveCategoryOrder,
  deleteCategory: deleteCategory,
  addTransaction: addTransaction,
  addMultipleTransactions: addMultipleTransactions,
  updateTransaction: updateTransaction,
  deleteTransaction: deleteTransaction,
  setBudget: setBudget,
  setOpening: setOpening,
  setTarget: setTarget,
  exportAllData: exportAllData,
  scanReceipt: scanReceipt,
  parseVoiceTransaction: parseVoiceTransaction,
  saveEvent: saveEvent,
  toggleEventStatus: toggleEventStatus,
  deleteEvent: deleteEvent,
  getEventTransactions: getEventTransactions,
  getLiveExchangeRate: getLiveExchangeRate,
  analyzeMonthWithAI: analyzeMonthWithAI
};

function doPost(e) {
  var out;
  try {
    var req = JSON.parse(e.postData.contents);
    var token = PropertiesService.getScriptProperties().getProperty('API_TOKEN');
    if (!token || req.token !== token) throw new Error('ERR_UNAUTHORIZED');
    var fn = API_MAP_[req.fn];
    if (!fn) throw new Error('ERR_UNKNOWN_FN');
    out = { ok: true, data: fn.apply(null, req.args || []) };
  } catch (err) {
    out = { ok: false, error: String(err && err.message ? err.message : err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

function tesUkuranCache() {
  var d = getData('2026-11');
  Logger.log('Ukuran: ' + JSON.stringify(d).length + ' karakter');
}
