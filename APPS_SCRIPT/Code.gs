// AT FARM V10.93 - FIX RENDER TABEL TOTAL KEUANGAN
// AT FARM V10.68 REKAP KEUANGAN TERPUSAT
// AT FARM V10.67 - FIX PEMBACA DATA LANGSUNG
// AT FARM V10.66 - LIST DI LUAR BOX
// AT FARM V10.65 - DAFTAR BARU UANG MASUK/KELUAR
// AT FARM V10.63 FIX FINAL DATA UANG MASUK & PENGELUARAN
// AT FARM V10.56 - LAPORAN UANG MASUK PENGELUARAN; BASE V10.53
// AT FARM V10.53 - FIX KEUANGAN GABUNGAN + TOTAL LABA
const CONFIG = {
  SPREADSHEET_ID: '1iH-U2bS3TtoxedqNfdh0Yh1GIUokaHCakNbboyPGLyI',
  MASTER_SHEET: 'DATA',
  INVENTORY_SHEET: 'INVENTORI',
  CASHIER_MASTER_SHEET: 'DATA2',
  CASHIER_TX_SHEET: 'KASIR_TRANSAKSI',
  EXPENSE_SHEET: 'PENGELUARAN',
  USERS_SHEET: 'USERS',
  HEADER_ROW: 1,
  DATA_START_ROW: 2,
  MAX_RECENT_ROWS: 100000,
  DEFAULT_SHEETS: ['KRATON', 'KAUMAN', 'RANGGEH']
};

const TX_HEADERS = [
  'No', 'Tanggal', 'Sayur', 'Jumlah', 'Satuan',
  'Harga Beli', 'PHP Harga Jual', 'Total Penjualan', 'Laba', 'Jumlah Dibayar', 'Status'
];

function getSpreadsheet_() {
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}

function doGet() {
  // V10.85: data REKAP_KEUANGAN A:F ditempel langsung ke HTML saat halaman dibuka.
  // Tidak memakai google.script.run untuk list rekap, sehingga tidak ada spinner yang menunggu request kedua.
  const tpl = HtmlService.createTemplateFromFile('Index');
  return tpl.evaluate()
    .setTitle('Aplikasi Penjualan Sayur & Buah')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🥦 Input Data Sayuran')
    .addItem('Buka Form Input', 'showSidebar')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createTemplateFromFile('Sidebar')
    .evaluate()
    .setTitle('Input Data Sayuran');
  SpreadsheetApp.getUi().showSidebar(html);
}

function getInitialData() {
  const ss = getSpreadsheet_();
  const dataSheet = getOrCreateDataSheet_();
  ensureDataHeaders_(dataSheet);

  // Pastikan semua sheet transaksi sudah menggunakan format terbaru.
  ss.getSheets().forEach(sheet => {
    const name = sheet.getName();
    if (name !== CONFIG.MASTER_SHEET && name !== CONFIG.INVENTORY_SHEET && name !== CONFIG.CASHIER_MASTER_SHEET && name !== CONFIG.CASHIER_TX_SHEET && name !== CONFIG.EXPENSE_SHEET && name !== 'MEMBER' && name !== CONFIG.USERS_SHEET && name !== 'REKAP_KEUANGAN' && name !== 'UANG_MASUK') {
      ensureTransactionHeaders_(sheet);
    }
  });

  const sheets = getTransactionSheetNames_();
  const products = getProductsFromData_();
  const globalPhp = getGlobalPHP_();
  const satuan = [...new Set(products.map(p => p.unit).filter(Boolean))].sort();

  return {
    sheets: sheets,
    vegetables: products,
    satuanList: satuan,
    globalPHP: globalPhp,
    defaultSheets: CONFIG.DEFAULT_SHEETS.filter(x => sheets.indexOf(x) >= 0)
  };
}

function getTransactionSheetNames_() {
  return getSpreadsheet_().getSheets()
    .map(s => s.getName())
    .filter(name => name !== CONFIG.MASTER_SHEET && name !== CONFIG.INVENTORY_SHEET && name !== CONFIG.CASHIER_MASTER_SHEET && name !== CONFIG.CASHIER_TX_SHEET && name !== 'MEMBER' && name !== CONFIG.USERS_SHEET && name !== CONFIG.EXPENSE_SHEET && name !== 'REKAP_KEUANGAN' && name !== 'UANG_MASUK' && name !== AT_SESSION_SHEET_V1111);
}

function getOrCreateDataSheet_() {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(CONFIG.MASTER_SHEET);
  if (!sheet) sheet = ss.insertSheet(CONFIG.MASTER_SHEET);
  return sheet;
}

function ensureDataHeaders_(sheet) {
  if (sheet.getMaxColumns() < 4) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), 4 - sheet.getMaxColumns());
  }
  const headers = sheet.getRange(1, 1, 1, 4).getValues()[0];
  if (!headers[0]) sheet.getRange(1, 1).setValue('No');
  if (!headers[1]) sheet.getRange(1, 2).setValue('Nama Sayur');
  if (!headers[2]) sheet.getRange(1, 3).setValue('Satuan');
  if (!headers[3] || normalizeKey_(headers[3]) !== 'php harga jual') {
    sheet.getRange(1, 4).setValue('PHP Harga Jual');
  }
  sheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#e8f5e9');
}

function getProductsFromData_() {
  const sheet = getOrCreateDataSheet_();
  ensureDataHeaders_(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  const map = {};

  values.forEach((r, i) => {
    const name = String(r[1] || '').trim();
    if (!name) return;
    const key = normalizeKey_(name);
    map[key] = {
      name: name,
      unit: String(r[2] || '').trim(),
      php: number_(r[3]),
      dataRow: i + 2
    };
  });

  return Object.keys(map)
    .map(k => map[k])
    .sort((a, b) => a.name.localeCompare(b.name, 'id'));
}

/**
 * PHP adalah nilai GLOBAL.
 * Disimpan di DATA kolom D dan diterapkan ke semua barang.
 * Untuk kompatibilitas dengan data lama, jika D2 kosong akan dicari
 * nilai PHP pertama yang tersedia lalu disalin ke seluruh kolom D.
 */
function getGlobalPHP_() {
  const sheet = getOrCreateDataSheet_();
  ensureDataHeaders_(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  const values = sheet.getRange(2, 4, lastRow - 1, 1).getValues().flat();
  let global = number_(values[0]);

  if (!global) {
    for (let i = 0; i < values.length; i++) {
      const n = number_(values[i]);
      if (n > 0) {
        global = n;
        break;
      }
    }
  }

  // Samakan D seluruh barang agar PHP benar-benar global.
  if (global >= 0 && lastRow >= 2) {
    const normalized = values.map(() => [global]);
    sheet.getRange(2, 4, normalized.length, 1).setValues(normalized);
    sheet.getRange(2, 4, normalized.length, 1).setNumberFormat('#,##0');
  }

  return global;
}

function setGlobalPHP_(php) {
  const value = number_(php);
  if (value < 0) throw new Error('PHP tidak boleh negatif.');

  const sheet = getOrCreateDataSheet_();
  ensureDataHeaders_(sheet);
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    // Belum ada produk: simpan PHP global di D2 tanpa membuat barang palsu.
    sheet.getRange(2, 4).setValue(value);
    sheet.getRange(2, 4).setNumberFormat('#,##0');
    return value;
  }

  const values = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  const out = values.map(r => [r[0], r[1], r[2], value]);
  sheet.getRange(2, 1, out.length, 4).setValues(out);
  sheet.getRange(2, 4, out.length, 1).setNumberFormat('#,##0');
  return value;
}

function getSheetData(sheetName, startDate, endDate) {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(String(sheetName || '').trim());
  if (!sheet || sheet.getName() === CONFIG.MASTER_SHEET || sheet.getName() === CONFIG.INVENTORY_SHEET) {
    return [];
  }

  ensureTransactionHeaders_(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  // Baca seluruh kolom transaksi yang tersedia setelah migrasi.
  const values = sheet.getRange(2, 1, lastRow - 1, TX_HEADERS.length).getValues();
  const rows = [];

  values.forEach((r, i) => {
    if (!r[1] && !r[2]) return;
    rows.push(normalizeTransactionRow_(r, i + 2));
  });

  // Filter tanggal untuk Data Terbaru.
  const start = startDate ? startOfDay_(parseDateInput(startDate)) : null;
  const end = endDate ? endOfDay_(parseDateInput(endDate)) : null;
  const filteredRows = rows.filter(r => {
    if (!r.tanggalMs) return !start && !end;
    if (start && r.tanggalMs < start.getTime()) return false;
    if (end && r.tanggalMs > end.getTime()) return false;
    return true;
  });

  // Urutan: tanggal lama -> tanggal baru. Jika tanggal sama, No kecil -> besar.
  filteredRows.sort((a, b) => {
    const da = Number(a.tanggalMs || 0);
    const db = Number(b.tanggalMs || 0);
    if (da !== db) return da - db;
    return Number(a.no || 0) - Number(b.no || 0);
  });

  // Ambil maksimal data terakhir, tetapi tetap tampil ASCENDING.
  return filteredRows.slice(Math.max(0, filteredRows.length - CONFIG.MAX_RECENT_ROWS));
}

function getRecentData(sheetFilter, startDate, endDate) {
  const filter = String(sheetFilter || 'SEMUA SHEET').trim();
  const sheetNames = filter === 'SEMUA SHEET' ? getTransactionSheetNames_() : [filter];
  let rows = [];

  sheetNames.forEach(name => {
    const data = getSheetData(name, startDate, endDate);
    data.forEach(row => {
      row.sheetName = name;
      rows.push(row);
    });
  });

  rows.sort((a, b) => {
    const da = a.tanggalInput ? new Date(a.tanggalInput + 'T00:00:00').getTime() : 0;
    const db = b.tanggalInput ? new Date(b.tanggalInput + 'T00:00:00').getTime() : 0;
    if (da !== db) return da - db;
    const na = Number(a.no || 0), nb = Number(b.no || 0);
    if (na !== nb) return na - nb;
    return String(a.sheetName || '').localeCompare(String(b.sheetName || ''));
  });

  return rows.slice(Math.max(0, rows.length - CONFIG.MAX_RECENT_ROWS));
}

function normalizeTransactionRow_(r, rowNumber) {
  const tanggalObj = parseDateValue_(r[1]);
  const jumlah = number_(r[3]);
  const hargaBeli = number_(r[5]);
  const php = number_(r[6]);
  const totalPenjualan = number_(r[7]) || ((hargaBeli + php) * jumlah);
  const laba = number_(r[8]) || (php * jumlah);

  return {
    rowNumber: rowNumber,
    no: r[0],
    tanggalMs: tanggalObj ? tanggalObj.getTime() : 0,
    tanggal: tanggalObj ? formatDate_(tanggalObj) : String(r[1] || ''),
    tanggalInput: tanggalObj ? formatDateForInput_(tanggalObj) : '',
    tanggalHari: tanggalObj ? formatDateWithDay_(tanggalObj) : String(r[1] || ''),
    sayur: String(r[2] || ''),
    jumlah: jumlah,
    satuan: String(r[4] || ''),
    hargaBeli: hargaBeli,
    php: php,
    hargaJual: hargaBeli + php,
    totalPenjualan: totalPenjualan,
    totalJual: totalPenjualan,
    total: totalPenjualan,
    laba: laba,
    jumlahDibayar: number_(r[9]),
    status: normalizeStatus_(r[10])
  };
}

function ensureTransactionHeaders_(sheet) {
  let maxCols = sheet.getMaxColumns();
  const lastColRead = Math.min(Math.max(maxCols, 1), 11);
  let h = sheet.getRange(1, 1, 1, lastColRead).getValues()[0].map(v => normalizeKey_(v));

  // Format v4: 10 kolom dengan Status di kolom J. Tambahkan kolom Jumlah Dibayar
  // sebelum Status agar struktur baru menjadi 11 kolom.
  const hasOldStatus10 = h.length >= 10 && h[9] === 'status' && h[0] === 'no' && h[1] === 'tanggal';
  if (hasOldStatus10 && h.length < 11) {
    sheet.insertColumnBefore(10);
    maxCols = sheet.getMaxColumns();
    h = sheet.getRange(1, 1, 1, Math.min(maxCols, 11)).getValues()[0].map(v => normalizeKey_(v));
    // Kolom J baru kosong; Status lama otomatis bergeser ke K.
    sheet.getRange(1, 10).setValue('Jumlah Dibayar');
    sheet.getRange(1, 11).setValue('Status');
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      sheet.getRange(2, 10, lastRow - 1, 1).clearContent();
    }
  }

  // Format lama 10 kolom tanpa status.
  h = sheet.getRange(1, 1, 1, Math.min(sheet.getMaxColumns(), 11)).getValues()[0].map(v => normalizeKey_(v));
  const old10 = h.length >= 10 &&
    h[0] === 'no' && h[1] === 'tanggal' && h[2] === 'sayur' &&
    h[3] === 'jumlah' && h[4] === 'satuan' && h[5] === 'harga beli' &&
    h[6] === 'harga jual' && h[7].indexOf('total beli') >= 0 && h[8].indexOf('total jual') >= 0;

  const old7 = h.length >= 7 &&
    h[0] === 'no' && h[1] === 'tanggal' && h[2] === 'sayur' &&
    h[3] === 'jumlah' && h[4] === 'satuan' &&
    (h[5] === 'harga' || h[5] === 'harga jual' || h[6] === 'total');

  if (old10) {
    migrateOldTenColumnSheet_(sheet);
  } else if (old7) {
    migrateOldSevenColumnSheet_(sheet);
  }

  if (sheet.getMaxColumns() < TX_HEADERS.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), TX_HEADERS.length - sheet.getMaxColumns());
  }

  sheet.getRange(1, 1, 1, TX_HEADERS.length).setValues([TX_HEADERS]);
  sheet.getRange(1, 1, 1, TX_HEADERS.length)
    .setFontWeight('bold')
    .setBackground('#e8f5e9');

  // Data lama tanpa status dianggap LUNAS. DP lama tanpa jumlah dibayar = 0.
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const statusRange = sheet.getRange(2, 11, lastRow - 1, 1);
    const statuses = statusRange.getValues().map(r => [normalizeStatus_(r[0])]);
    statusRange.setValues(statuses);
    sheet.getRange(2, 10, lastRow - 1, 1).setNumberFormat('#,##0');
  }
  sheet.setFrozenRows(1);
}

function migrateOldSevenColumnSheet_(sheet) {
  const lastRow = sheet.getLastRow();
  if (sheet.getMaxColumns() < 11) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), 11 - sheet.getMaxColumns());
  }
  if (lastRow < 2) {
    sheet.getRange(1, 1, 1, 11).setValues([TX_HEADERS]);
    return;
  }

  const rows = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
  const converted = rows.map(r => {
    const jumlah = number_(r[3]);
    const oldHarga = number_(r[5]);
    const oldTotal = number_(r[6]) || (oldHarga * jumlah);
    return [r[0], r[1], r[2], r[3], r[4], 0, oldHarga, oldTotal, oldTotal, oldTotal, 'LUNAS'];
  });

  sheet.getRange(2, 1, converted.length, 11).setValues(converted);
  sheet.getRange(1, 1, 1, 11).setValues([TX_HEADERS]);
}

function migrateOldTenColumnSheet_(sheet) {
  const lastRow = sheet.getLastRow();
  if (sheet.getMaxColumns() < 11) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), 11 - sheet.getMaxColumns());
  }
  if (lastRow < 2) {
    sheet.getRange(1, 1, 1, 11).setValues([TX_HEADERS]);
    return;
  }

  const rows = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
  const converted = rows.map(r => {
    const beli = number_(r[5]);
    const jualLama = number_(r[6]);
    const totalJualLama = number_(r[8]);
    const jumlah = number_(r[3]);
    const php = Math.max(0, jualLama - beli);
    const totalPenjualan = totalJualLama || ((beli + php) * jumlah);
    const laba = number_(r[9]) || (php * jumlah);
    return [r[0], r[1], r[2], r[3], r[4], beli, php, totalPenjualan, laba, totalPenjualan, 'LUNAS'];
  });

  sheet.getRange(2, 1, converted.length, 11).setValues(converted);
  sheet.getRange(1, 1, 1, 11).setValues([TX_HEADERS]);
}

function getSummaryAllSheets(sheetName, startDate, endDate) {
  return buildSummary_(sheetName || '', startDate || '', endDate || '', null);
}

function getSummaryByDate(startDate, endDate) {
  return buildSummary_('', startDate || '', endDate || '', null);
}

function getSummaryBySheetDate(sheetName, startDate, endDate) {
  return buildSummary_(sheetName || '', startDate || '', endDate || '', null);
}

// Rekap Laporan: SEMUA sheet yang berasal dari FORM INPUT PENJUALAN.
// Sheet sistem (DATA, INVENTORI, DATA2, KASIR_TRANSAKSI, MEMBER) otomatis dikecualikan
// oleh getTransactionSheetNames_(). Lokasi baru yang dibuat dari Form Penjualan otomatis ikut.
function getSummaryUtama(sheetName, startDate, endDate) {
  return buildSummary_(sheetName || '', startDate || '', endDate || '', null);
}

// Rekap lokasi lain: semua sheet selain KRATON, KAUMAN, RANGGEH,
// DATA dan INVENTORI.
function getSummaryLuar(sheetName, startDate, endDate) {
  const outside = getTransactionSheetNames_().filter(name => CONFIG.DEFAULT_SHEETS.indexOf(name) < 0);
  return buildSummary_(sheetName || '', startDate || '', endDate || '', outside);
}

function buildSummary_(sheetName, startDate, endDate, allowedSheets) {
  const ss = getSpreadsheet_();
  const normalizedSheet = normalizeKey_(sheetName);
  const allSheets = !sheetName || normalizedSheet === 'semua sheet' ||
    normalizedSheet === 'semua' || normalizedSheet === 'semua sheet utama' ||
    normalizedSheet === 'semua lokasi luar';

  let names;
  if (allowedSheets) {
    names = allSheets ? allowedSheets.slice() :
      (allowedSheets.map(normalizeKey_).indexOf(normalizedSheet) >= 0 ? [sheetName] : []);
  } else {
    names = allSheets ? getTransactionSheetNames_() : [sheetName];
  }

  const start = startDate ? startOfDay_(parseDateInput(startDate)) : null;
  const end = endDate ? endOfDay_(parseDateInput(endDate)) : null;
  const map = {};
  let grandTotal = 0;
  let grandProfit = 0;
  let grandLunas = 0;
  let grandUtang = 0;
  let grandDP = 0;

  names.forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet || name === CONFIG.MASTER_SHEET || name === CONFIG.INVENTORY_SHEET) return;

    ensureTransactionHeaders_(sheet);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    const rows = sheet.getRange(2, 1, lastRow - 1, TX_HEADERS.length).getValues();
    rows.forEach(r => {
      if (!r[1] && !r[2]) return;
      const d = parseDateValue_(r[1]);
      if (!d) return;
      if (start && d < start) return;
      if (end && d > end) return;

      const sayur = String(r[2] || '').trim();
      const satuan = String(r[4] || '').trim();
      const jumlah = number_(r[3]);
      const beli = number_(r[5]);
      const php = number_(r[6]);
      const total = number_(r[7]) || ((beli + php) * jumlah);
      const laba = number_(r[8]) || (php * jumlah);
      const status = normalizeStatus_(r[10]);
      const key = normalizeKey_(sayur) + '|' + normalizeKey_(satuan);

      if (!map[key]) map[key] = { sayur, satuan, jumlah: 0, total: 0, laba: 0 };
      map[key].jumlah += jumlah;
      map[key].total += total;
      map[key].laba += laba;
      grandTotal += total;
      grandProfit += laba;
      if (status === 'LUNAS') grandLunas += total;
      else if (status === 'UTANG') grandUtang += total;
      else if (status === 'DP') grandDP += total;
    });
  });

  const data = Object.keys(map).map(k => map[k]);
  data.sort((a, b) => {
    const order = unitOrder_(a.satuan) - unitOrder_(b.satuan);
    if (order !== 0) return order;
    return a.sayur.localeCompare(b.sayur, 'id');
  });

  return {
    data: data,
    grandTotal: grandTotal,
    grandProfit: grandProfit,
    grandModal: grandTotal - grandProfit,
    grandLunas: grandLunas,
    grandUtang: grandUtang,
    grandDP: grandDP
  };
}

// V11.11: metadata terpisah. Tidak mengubah transaksi atau baris historis.
// Satu pemanggilan simpan = satu sesi; ID sesi yang sama boleh dipakai bila
// kelak sebuah proses simpan mengirim beberapa baris sekaligus.
const AT_SESSION_SHEET_V1111 = 'AT_SESI_KEUANGAN';
function atSessionSheetV1111_() {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(AT_SESSION_SHEET_V1111);
  if (!sh) {
    sh = ss.insertSheet(AT_SESSION_SHEET_V1111);
    sh.getRange(1,1,1,5).setValues([['SHEET','BARIS','SESI','SUMBER','WAKTU CATAT']]);
    sh.setFrozenRows(1);
  }
  return sh;
}
function atRecordSessionV1111_(sheetName,rowNumber,source,sessionId) {
  const sh = atSessionSheetV1111_();
  sh.appendRow([sheetName,rowNumber,String(sessionId||Utilities.getUuid()),source,new Date()]);
}
function atReadSessionsV1111_() {
  const sh = getSpreadsheet_().getSheetByName(AT_SESSION_SHEET_V1111);
  const map = {};
  if (!sh || sh.getLastRow()<2) return map;
  sh.getRange(2,1,sh.getLastRow()-1,4).getValues().forEach(function(r){
    map[String(r[0])+'|'+String(r[1])] = {session:String(r[2]),source:String(r[3])};
  });
  return map;
}
function atGroupNewSessionRowsV1111_(rows) {
  const grouped=[], byKey={};
  rows.forEach(function(r){
    if (!r._atSession) { grouped.push(r); return; }
    const key = [r._atSession,r.tanggal,r._atSource||r.sumber,r.jenis].join('|');
    if (!byKey[key]) {
      const item=Object.assign({},r);
      item.keterangan=String(r._atSource||r.sumber);
      byKey[key]=item; grouped.push(item);
    } else {
      byKey[key].uangMasuk += number_(r.uangMasuk);
      byKey[key].uangKeluar += number_(r.uangKeluar);
    }
  });
  return grouped;
}

function simpanDataSayur(payload) {
  validateTransaction_(payload);
  const sheet = getOrCreateTransactionSheet_(payload.targetSheet);
  const dataSheet = getOrCreateDataSheet_();
  ensureDataHeaders_(dataSheet);

  // Selalu gunakan PHP global terbaru dari DATA.
  const globalPhp = getGlobalPHP_();
  const tanggal = parseDateInput(payload.tanggal);
  const jumlah = number_(payload.jumlah);
  const hargaBeli = number_(payload.hargaBeli);
  const php = globalPhp;
  const total = (hargaBeli + php) * jumlah;
  const laba = php * jumlah;
  const no = getNextNumber(payload.targetSheet);

  sheet.appendRow([
    no, tanggal, payload.sayur.trim(), jumlah, payload.satuan.trim(),
    hargaBeli, php, total, laba, number_(payload.dpAmount), normalizeStatus_(payload.status)
  ]);
  // Catat hanya transaksi baru; bila gagal, transaksi tetap tersimpan normal.
  try { atRecordSessionV1111_(sheet.getName(),sheet.getLastRow(),sheet.getName(),payload.sessionId); }
  catch(e) { console.log('Metadata sesi AT FARM:',e); }

  ensureProductInData_(payload.sayur.trim(), payload.satuan.trim(), php);
  formatTransactionSheet_(sheet);

  // V10.89: setiap transaksi penjualan langsung memperbarui pusat data keuangan.
  try { syncRekapKeuangan_(); } catch (e) { console.log('Sync REKAP_KEUANGAN:', e); }

  return { success: true, no: no, total: total, laba: laba, php: php };
}

function validateTransaction_(p) {
  if (!p) throw new Error('Data tidak ditemukan.');
  if (!p.targetSheet) throw new Error('Pilih Sheet/Lokasi.');
  if (!p.tanggal) throw new Error('Tanggal wajib diisi.');
  if (!p.sayur || !String(p.sayur).trim()) throw new Error('Nama sayur wajib diisi.');
  if (number_(p.jumlah) <= 0) throw new Error('Jumlah harus lebih dari 0.');
  if (!p.satuan || !String(p.satuan).trim()) throw new Error('Satuan wajib dipilih.');
  if (number_(p.hargaBeli) < 0) throw new Error('Harga beli tidak boleh negatif.');
  const status = normalizeStatus_(p.status);
  const dp = number_(p.dpAmount);
  if (status === 'DP' && dp <= 0) throw new Error('Jika memilih DP, jumlah yang dibayar wajib diisi.');
  if (status === 'DP' && dp > ((number_(p.hargaBeli) + getGlobalPHP_()) * number_(p.jumlah))) throw new Error('Jumlah DP tidak boleh lebih besar dari total penjualan.');
}

function getOrCreateTransactionSheet_(name) {
  const clean = sanitizeSheetName_(name);
  if (!clean || clean === CONFIG.MASTER_SHEET || clean === CONFIG.INVENTORY_SHEET) {
    throw new Error('Nama sheet tidak valid.');
  }
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(clean);
  if (!sheet) sheet = ss.insertSheet(clean);
  ensureTransactionHeaders_(sheet);
  return sheet;
}

function createNewLocationSheet(locationName) {
  const clean = sanitizeSheetName_(locationName);
  if (!clean) throw new Error('Nama lokasi tidak valid.');
  if (clean === CONFIG.MASTER_SHEET || clean === CONFIG.INVENTORY_SHEET) {
    throw new Error('Nama tersebut digunakan untuk sheet sistem.');
  }

  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(clean);
  if (!sheet) {
    sheet = ss.insertSheet(clean);
    ensureTransactionHeaders_(sheet);
    formatTransactionSheet_(sheet);
  }
  return { success: true, sheetName: sheet.getName() };
}

function getNextNumber(sheetName) {
  const sheet = getOrCreateTransactionSheet_(sheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 1;

  const vals = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  let max = 0;
  vals.forEach(v => {
    const n = Number(v);
    if (isFinite(n) && n > max) max = n;
  });
  return max + 1;
}

function getDataByRow(sheetName, rowNumber) {
  const sheet = getSpreadsheet_().getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet tidak ditemukan.');
  ensureTransactionHeaders_(sheet);
  const r = sheet.getRange(rowNumber, 1, 1, TX_HEADERS.length).getValues()[0];
  return normalizeTransactionRow_(r, rowNumber);
}

function updateDataSayur(payload) {
  validateTransaction_(payload);
  const sheet = getSpreadsheet_().getSheetByName(payload.targetSheet);
  if (!sheet) throw new Error('Sheet tidak ditemukan.');
  ensureTransactionHeaders_(sheet);

  const rowNumber = Number(payload.rowNumber);
  if (rowNumber < 2 || rowNumber > sheet.getLastRow()) throw new Error('Baris data tidak valid.');

  const jumlah = number_(payload.jumlah);
  const hargaBeli = number_(payload.hargaBeli);
  const php = getGlobalPHP_();
  const total = (hargaBeli + php) * jumlah;
  const laba = php * jumlah;

  sheet.getRange(rowNumber, 1, 1, TX_HEADERS.length).setValues([[
    Number(payload.no), parseDateInput(payload.tanggal), payload.sayur.trim(), jumlah,
    payload.satuan.trim(), hargaBeli, php, total, laba, number_(payload.dpAmount), normalizeStatus_(payload.status)
  ]]);

  ensureProductInData_(payload.sayur.trim(), payload.satuan.trim(), php);
  formatTransactionSheet_(sheet);
  return { success: true, total: total, laba: laba, php: php };
}

function updateStatusPembayaran(sheetName, rowNumber, status) {
  const sheet = getSpreadsheet_().getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet tidak ditemukan.');
  ensureTransactionHeaders_(sheet);
  const row = Number(rowNumber);
  if (row < 2 || row > sheet.getLastRow()) throw new Error('Baris tidak valid.');

  const cleanStatus = normalizeStatus_(status);
  sheet.getRange(row, 11).setValue(cleanStatus);
  if (cleanStatus === 'LUNAS') {
    const total = number_(sheet.getRange(row, 8).getValue());
    sheet.getRange(row, 10).setValue(total);
  }
  return { success: true, status: cleanStatus };
}

function tandaiLunas(sheetName, rowNumber) {
  return updateStatusPembayaran(sheetName, rowNumber, 'LUNAS');
}

function hapusDataSayur(sheetName, rowNumber) {
  const sheet = getSpreadsheet_().getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet tidak ditemukan.');
  if (Number(rowNumber) < 2 || Number(rowNumber) > sheet.getLastRow()) throw new Error('Baris tidak valid.');
  sheet.deleteRow(Number(rowNumber));
  return { success: true };
}

/**
 * Edit PHP GLOBAL. Tidak lagi bergantung pada barang yang sedang dipilih.
 * Satu kali edit akan mengubah PHP untuk seluruh barang di DATA.
 */
function updatePHP(php) {
  const value = number_(php);
  if (value < 0) throw new Error('PHP tidak boleh negatif.');

  setGlobalPHP_(value);

  return {
    success: true,
    php: value,
    vegetables: getProductsFromData_()
  };
}

function ensureProductInData_(sayur, satuan, php) {
  const sheet = getOrCreateDataSheet_();
  ensureDataHeaders_(sheet);
  const lastRow = sheet.getLastRow();
  let found = -1;

  if (lastRow >= 2) {
    const names = sheet.getRange(2, 2, lastRow - 1, 1).getValues().flat();
    names.forEach((v, i) => {
      if (normalizeKey_(v) === normalizeKey_(sayur)) found = i + 2;
    });
  }

  if (found === -1) {
    const nameValues = lastRow >= 2 ? sheet.getRange(2, 2, lastRow - 1, 1).getValues().flat() : [];
    let row = -1;
    for (let i = 0; i < nameValues.length; i++) {
      if (!String(nameValues[i] || '').trim()) {
        row = i + 2;
        break;
      }
    }
    if (row === -1) row = sheet.getLastRow() + 1;
    const globalPhp = getGlobalPHP_() || number_(php);
    sheet.getRange(row, 1, 1, 4).setValues([[row - 1, sayur, satuan, globalPhp]]);
  } else {
    if (satuan) sheet.getRange(found, 3).setValue(satuan);
    // Pastikan PHP barang tersebut sama dengan PHP global.
    sheet.getRange(found, 4).setValue(getGlobalPHP_());
  }
}

/**
 * Menambah/memperbarui nama sayur dan satuan pada Sheet DATA.
 * PHP tetap menggunakan PHP global yang tersimpan di DATA kolom D.
 */
function addProductToData(name, unit) {
  const cleanName = String(name || '').trim();
  const cleanUnit = String(unit || '').trim();
  if (!cleanName) throw new Error('Nama sayur/buah wajib diisi.');
  if (!cleanUnit) throw new Error('Satuan wajib diisi.');

  const sheet = getOrCreateDataSheet_();
  ensureDataHeaders_(sheet);
  const lastRow = sheet.getLastRow();
  const globalPhp = getGlobalPHP_();

  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
    for (let i = 0; i < values.length; i++) {
      if (normalizeKey_(values[i][1]) === normalizeKey_(cleanName)) {
        const row = i + 2;
        sheet.getRange(row, 3).setValue(cleanUnit);
        sheet.getRange(row, 4).setValue(globalPhp);
        return {success:true, action:'updated', row:row, name:cleanName, unit:cleanUnit, php:globalPhp};
      }
    }
  }

  let targetRow = -1;
  if (lastRow >= 2) {
    const names = sheet.getRange(2, 2, lastRow - 1, 1).getValues().flat();
    for (let i = 0; i < names.length; i++) {
      if (!String(names[i] || '').trim()) {
        targetRow = i + 2;
        break;
      }
    }
  }
  if (targetRow === -1) targetRow = Math.max(2, sheet.getLastRow() + 1);

  const no = targetRow - 1;
  sheet.getRange(targetRow, 1, 1, 4).setValues([[no, cleanName, cleanUnit, globalPhp]]);
  sheet.getRange(targetRow, 4).setNumberFormat('#,##0');
  return {success:true, action:'added', row:targetRow, name:cleanName, unit:cleanUnit, php:globalPhp};
}

function parseDateInput(value) {
  if (value instanceof Date && !isNaN(value)) return value;
  const s = String(value || '').trim();
  if (!s) throw new Error('Tanggal tidak valid.');

  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));

  const d = new Date(s);
  if (isNaN(d)) throw new Error('Tanggal tidak valid.');
  return d;
}

function parseDateValue_(value) {
  if (value instanceof Date && !isNaN(value)) return new Date(value.getTime());
  if (!value) return null;
  const d = new Date(value);
  return isNaN(d) ? null : d;
}

function formatDate_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'dd/MM/yyyy');
}

function formatDateForInput_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function formatDateWithDay_(date) {
  const days = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
  const d = new Date(date.getTime());
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'd/MM/yyyy') + ' (' + days[d.getDay()] + ')';
}

function getDayName_(date) {
  const days = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
  return days[date.getDay()];
}

function formatPeriod_(startDate, endDate) {
  if (!startDate && !endDate) return 'Semua Periode';

  if (startDate && endDate) {
    const s = startOfDay_(parseDateInput(startDate));
    const e = startOfDay_(parseDateInput(endDate));
    const days = [];
    for (let d = new Date(s); d <= e; d.setDate(d.getDate() + 1)) {
      days.push(getDayName_(new Date(d)));
    }
    return formatDateShort_(s) + ' s/d ' + formatDateShort_(e) + ' (' + days.join(',') + ')';
  }

  if (startDate) {
    const s = parseDateInput(startDate);
    return 'Mulai ' + formatDateShort_(s) + ' (' + getDayName_(s) + ')';
  }

  const e = parseDateInput(endDate);
  return 'Sampai ' + formatDateShort_(e) + ' (' + getDayName_(e) + ')';
}

function getPrintMeta(sheetName, startDate, endDate) {
  return {
    sheet: sheetName || 'Semua Sheet',
    period: formatPeriod_(startDate, endDate),
    printedAt: formatDateWithTime_(new Date())
  };
}

function formatDateShort_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'd/MM/yyyy');
}

function formatDateWithTime_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'd/MM/yyyy HH:mm');
}

function startOfDay_(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate(), 0, 0, 0, 0);
}

function endOfDay_(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate(), 23, 59, 59, 999);
}

function normalizeStatus_(v) {
  const s = String(v == null ? '' : v).trim().toUpperCase();
  if (s === 'UTANG') return 'UTANG';
  if (s === 'DP') return 'DP';
  return 'LUNAS';
}

function normalizeKey_(v) {
  return String(v || '').trim().toLowerCase();
}

function number_(v) {
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  let s = String(v == null ? '' : v).trim();
  if (!s) return 0;

  s = s.replace(/[^0-9,.-]/g, '');
  // Mendukung format Indonesia: 1.500.000 / 1.500 / 12,5
  if (s.indexOf(',') >= 0 && s.indexOf('.') >= 0) {
    s = s.replace(/\./g, '').replace(',', '.');
  } else if (s.indexOf(',') >= 0) {
    s = s.replace(',', '.');
  } else if ((s.match(/\./g) || []).length > 1) {
    s = s.replace(/\./g, '');
  }

  const n = Number(s);
  return isFinite(n) ? n : 0;
}

function unitOrder_(u) {
  const s = normalizeKey_(u);
  if (s === 'kg') return 1;
  if (s === 'ikat') return 2;
  return 3;
}

function sanitizeSheetName_(name) {
  return String(name || '').trim().replace(/[\\/?*\[\]:]/g, '-').substring(0, 99);
}

function formatTransactionSheet_(sheet) {
  ensureTransactionHeaders_(sheet);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, TX_HEADERS.length)
    .setFontWeight('bold')
    .setBackground('#e8f5e9');

  if (sheet.getLastRow() >= 2) {
    sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).setNumberFormat('dd/MM/yyyy');
    sheet.getRange(2, 6, sheet.getLastRow() - 1, 5).setNumberFormat('#,##0');
  }
}

// ==========================================================
// REKAP KEUANGAN + DATA DIAGRAM OMZET HARIAN
// ==========================================================
function getFinancialRecapByMode_(mode, sheetName, startDate, endDate) {
  const filter = String(sheetName || 'SEMUA SHEET').trim();
  const normalSheets = getTransactionSheetNames_();
  let names;
  if (mode === 'CASHIER') names = [CONFIG.CASHIER_TX_SHEET];
  else if (mode === 'COMBINED') names = normalSheets.concat([CONFIG.CASHIER_TX_SHEET]);
  else names = filter === 'SEMUA SHEET' ? normalSheets : [filter];
  const start = startDate ? startOfDay_(parseDateInput(startDate)) : null;
  const end = endDate ? endOfDay_(parseDateInput(endDate)) : null;
  const dailyMap = {};
  let totalOmzet=0,totalLaba=0,totalModal=0,totalPaid=0,totalReceivable=0;

  names.forEach(name => {
    // KASIR_TRANSAKSI mempunyai struktur kolom khusus.
    if (name === CONFIG.CASHIER_TX_SHEET) {
      const sh=getOrCreateCashierTx_(), last=sh.getLastRow();
      if(last<2)return;
      const rows=sh.getRange(2,1,last-1,CASHIER_TX_HEADERS.length).getValues();
      const tx={};
      rows.forEach(r=>{
        const d=parseDateValue_(r[1]); if(!d)return;
        const day=startOfDay_(d); if(start&&day<start)return;if(end&&day>end)return;
        const id=String(r[0]||''); if(!tx[id])tx[id]={date:day,total:0,paid:number_(r[9]),status:normalizeStatus_(r[8]||'LUNAS')};
        tx[id].total+=number_(r[7]);
      });
      Object.keys(tx).forEach(id=>{
        const t=tx[id], omzet=t.total;
        const paid=t.status==='LUNAS'?omzet:(t.status==='DP'?Math.min(omzet,Math.max(0,t.paid)):0);
        const receivable=Math.max(0,omzet-paid);
        // DATA2 saat ini menyimpan harga jual Kasir, bukan HPP per transaksi,
        // sehingga laba Kasir tidak direka. Omzet/modal kasir tetap tercatat akurat.
        totalOmzet+=omzet; totalModal+=omzet; totalPaid+=paid; totalReceivable+=receivable;
        const key=Utilities.formatDate(t.date,Session.getScriptTimeZone(),'yyyy-MM-dd');
        if(!dailyMap[key])dailyMap[key]={date:key,label:Utilities.formatDate(t.date,Session.getScriptTimeZone(),'dd/MM'),omzet:0,laba:0};
        dailyMap[key].omzet+=omzet;
      });
      return;
    }

    const sheet=getSpreadsheet_().getSheetByName(name); if(!sheet)return;
    ensureTransactionHeaders_(sheet); const lastRow=sheet.getLastRow(); if(lastRow<2)return;
    sheet.getRange(2,1,lastRow-1,TX_HEADERS.length).getValues().forEach((r,i)=>{
      if(!r[1]&&!r[2])return; const row=normalizeTransactionRow_(r,i+2), d=parseDateValue_(r[1]); if(!d)return;
      const day=startOfDay_(d); if(start&&day<start)return;if(end&&day>end)return;
      const omzet=number_(row.totalPenjualan), laba=number_(row.laba), modal=omzet-laba;
      let paid=0;if(row.status==='LUNAS')paid=omzet;else if(row.status==='DP')paid=Math.min(omzet,Math.max(0,number_(row.jumlahDibayar)));
      const receivable=Math.max(0,omzet-paid);
      totalOmzet+=omzet;totalLaba+=laba;totalModal+=modal;totalPaid+=paid;totalReceivable+=receivable;
      const key=Utilities.formatDate(day,Session.getScriptTimeZone(),'yyyy-MM-dd');
      if(!dailyMap[key])dailyMap[key]={date:key,label:Utilities.formatDate(day,Session.getScriptTimeZone(),'dd/MM'),omzet:0,laba:0};
      dailyMap[key].omzet+=omzet;dailyMap[key].laba+=laba;
    });
  });
  return {daily:Object.keys(dailyMap).sort().map(k=>dailyMap[k]),totalOmzet,totalLaba,totalModal,totalPaid,totalReceivable};
}

function getSalesFinancialRecap(sheetName, startDate, endDate) {
  return getFinancialRecapByMode_('SALES', sheetName, startDate, endDate);
}
function getCashierFinancialRecap(startDate, endDate) {
  return getFinancialRecapByMode_('CASHIER', 'KASIR_TRANSAKSI', startDate, endDate);
}
function getCombinedFinancialRecap(startDate, endDate) {
  return getFinancialRecapByMode_('COMBINED', 'SEMUA SHEET', startDate, endDate);
}
// Kompatibilitas versi lama
function getFinancialRecap(sheetName, startDate, endDate) {
  return getSalesFinancialRecap(sheetName, startDate, endDate);
}

// ==========================================================
// KASIR / TRANSAKSI CEPAT - DATA2 TERPISAH DARI FORM UTAMA
// ==========================================================
const CASHIER_PRODUCT_HEADERS = ['No','Nama Produk','Satuan','Harga Dasar','URL Gambar','Aktif','PHP Kasir / Unit'];
const CASHIER_TX_HEADERS = ['ID Transaksi','Tanggal','Waktu','Nama Produk','Jumlah','Satuan','Harga','Subtotal','Status','Jumlah Dibayar','Member','Diskon'];
const MEMBER_SHEET = 'MEMBER';

function getOrCreateCashierMaster_(){
  const ss=getSpreadsheet_(); let sh=ss.getSheetByName(CONFIG.CASHIER_MASTER_SHEET);
  if(!sh) sh=ss.insertSheet(CONFIG.CASHIER_MASTER_SHEET);
  if(sh.getMaxColumns()<CASHIER_PRODUCT_HEADERS.length) sh.insertColumnsAfter(sh.getMaxColumns(),CASHIER_PRODUCT_HEADERS.length-sh.getMaxColumns());
  sh.getRange(1,1,1,CASHIER_PRODUCT_HEADERS.length).setValues([CASHIER_PRODUCT_HEADERS]).setFontWeight('bold').setBackground('#e8f5e9');
  return sh;
}
function getOrCreateCashierTx_(){
  const ss=getSpreadsheet_(); let sh=ss.getSheetByName(CONFIG.CASHIER_TX_SHEET);
  if(!sh) sh=ss.insertSheet(CONFIG.CASHIER_TX_SHEET);
  if(sh.getMaxColumns()<CASHIER_TX_HEADERS.length) sh.insertColumnsAfter(sh.getMaxColumns(),CASHIER_TX_HEADERS.length-sh.getMaxColumns());
  sh.getRange(1,1,1,CASHIER_TX_HEADERS.length).setValues([CASHIER_TX_HEADERS]).setFontWeight('bold').setBackground('#e8f5e9');
  sh.setFrozenRows(1); return sh;
}
function getCashierProducts(){
  const sh=getOrCreateCashierMaster_(), last=sh.getLastRow(); if(last<2)return [];
  return sh.getRange(2,1,last-1,CASHIER_PRODUCT_HEADERS.length).getValues().map((r,i)=>({row:i+2,no:r[0],name:String(r[1]||'').trim(),unit:String(r[2]||'').trim(),price:number_(r[3]),image:String(r[4]||'').trim(),active:String(r[5]||'YA').toUpperCase()!=='TIDAK',php:number_(r[6])})).filter(x=>x.name&&x.active);
}
function saveCashierProduct(p){
  p=p||{}; const name=String(p.name||'').trim(), unit=String(p.unit||'').trim(), price=number_(p.price), image=String(p.image||'').trim(), php=number_(p.php), requestedRow=Number(p.row||0);
  if(!name)throw new Error('Nama produk wajib diisi.'); if(!unit)throw new Error('Satuan wajib diisi.'); if(price<0)throw new Error('Harga tidak valid.'); if(php<0)throw new Error('PHP Kasir tidak valid.');
  const sh=getOrCreateCashierMaster_(), last=sh.getLastRow();
  // Saat mode edit, update berdasarkan nomor baris DATA2 agar nama produk boleh diubah dan foto lama tetap aman.
  if(requestedRow>=2 && requestedRow<=last){
    const oldImage=String(sh.getRange(requestedRow,5).getValue()||'').trim();
    const finalImage=image||oldImage;
    sh.getRange(requestedRow,2,1,6).setValues([[name,unit,price,finalImage,'YA',php]]);
    return {success:true,products:getCashierProducts()};
  }
  if(last>=2){const vals=sh.getRange(2,1,last-1,CASHIER_PRODUCT_HEADERS.length).getValues(); for(let i=0;i<vals.length;i++){if(normalizeKey_(vals[i][1])===normalizeKey_(name)){const oldImage=String(vals[i][4]||'').trim();sh.getRange(i+2,2,1,6).setValues([[name,unit,price,image||oldImage,'YA',php]]);return {success:true,products:getCashierProducts()};}}}
  sh.appendRow([Math.max(1,last),name,unit,price,image,'YA',php]); return {success:true,products:getCashierProducts()};
}
function getOrCreateMemberSheet_(){
  const ss=getSpreadsheet_(); let sh=ss.getSheetByName(MEMBER_SHEET);
  if(!sh) sh=ss.insertSheet(MEMBER_SHEET);
  sh.getRange(1,1,1,3).setValues([['No','Nama Member','Tanggal Daftar']]).setFontWeight('bold').setBackground('#e8f5e9');
  sh.setFrozenRows(1); return sh;
}
function getCashierMembers(){
  const sh=getOrCreateMemberSheet_(), last=sh.getLastRow(); if(last<2)return [];
  return sh.getRange(2,2,last-1,1).getDisplayValues().flat().map(x=>String(x||'').trim()).filter(Boolean).sort((a,b)=>a.localeCompare(b));
}
function saveMemberIfNeeded_(name){
  name=String(name||'').trim(); if(!name)return '';
  const sh=getOrCreateMemberSheet_(), last=sh.getLastRow();
  if(last>=2){const names=sh.getRange(2,2,last-1,1).getDisplayValues().flat();if(names.some(x=>normalizeKey_(x)===normalizeKey_(name)))return name;}
  sh.appendRow([Math.max(1,last),name,new Date()]); sh.getRange(sh.getLastRow(),3).setNumberFormat('dd/MM/yyyy'); return name;
}
function saveCashierTransaction(payload){
  payload=payload||{}; const items=payload.items||[]; if(!items.length)throw new Error('Keranjang masih kosong.');
  const status=normalizeStatus_(payload.status||'LUNAS'), paid=number_(payload.paid), member=saveMemberIfNeeded_(payload.member);
  const discountType=String(payload.discountType||'NOMINAL').toUpperCase(), discountInput=Math.max(0,number_(payload.discount));
  const now=new Date(), id=Utilities.formatDate(now,Session.getScriptTimeZone(),'yyyyMMdd-HHmmss')+'-'+Math.floor(Math.random()*900+100);
  const gross=items.reduce((a,it)=>a+number_(it.qty)*number_(it.price),0);
  const discount=Math.min(gross,discountType==='PERSEN'?gross*Math.min(discountInput,100)/100:discountInput);
  const total=Math.max(0,gross-discount), sh=getOrCreateCashierTx_();
  let allocated=0;
  const rows=items.map((it,idx)=>{const qty=number_(it.qty),price=number_(it.price),grossSub=qty*price;let share=idx===items.length-1?discount-allocated:(gross?discount*(grossSub/gross):0);allocated+=share;const netSub=Math.max(0,grossSub-share);return [id,now,Utilities.formatDate(now,Session.getScriptTimeZone(),'HH:mm:ss'),String(it.name||''),qty,String(it.unit||''),price,netSub,status,0,member,share];});
  const dibayar=status==='LUNAS'?total:(status==='DP'?Math.min(total,Math.max(0,paid)):0);
  rows.forEach(r=>r[9]=dibayar); const atStartRow=sh.getLastRow()+1; sh.getRange(atStartRow,1,rows.length,CASHIER_TX_HEADERS.length).setValues(rows); sh.getRange(2,2,sh.getLastRow()-1,1).setNumberFormat('dd/MM/yyyy');
  try { for(let k=0;k<rows.length;k++) atRecordSessionV1111_(sh.getName(),atStartRow+k,'KASIR',id); }
  catch(e) { console.log('Metadata sesi kasir:',e); }
  // V10.89: kasir ikut otomatis masuk REKAP_KEUANGAN.
  try { syncRekapKeuangan_(); } catch (e) { console.log('Sync REKAP_KEUANGAN:', e); }
  return {success:true,message:'Transaksi kasir berhasil disimpan.',id:id,total:total,discount:discount,member:member};
}
function getCashierRecap(startDate,endDate){
  const sh=getOrCreateCashierTx_(), last=sh.getLastRow(), start=startDate?startOfDay_(parseDateInput(startDate)):null, end=endDate?endOfDay_(parseDateInput(endDate)):null;
  const map={}; let grandTotal=0,totalPaid=0,totalReceivable=0; if(last<2)return {data:[],grandTotal:0,totalPaid:0,totalReceivable:0};
  sh.getRange(2,1,last-1,CASHIER_TX_HEADERS.length).getValues().forEach(r=>{const d=parseDateValue_(r[1]);if(!d)return;if(start&&d<start)return;if(end&&d>end)return;const name=String(r[3]||''),unit=String(r[5]||''),qty=number_(r[4]),sub=number_(r[7]);const key=normalizeKey_(name)+'|'+normalizeKey_(unit);if(!map[key])map[key]={sayur:name,satuan:unit,jumlah:0,total:0,laba:0};map[key].jumlah+=qty;map[key].total+=sub;grandTotal+=sub;});
  // jumlah dibayar disimpan per transaksi pada tiap item; hitung unik per ID
  const seen={}; sh.getRange(2,1,last-1,CASHIER_TX_HEADERS.length).getValues().forEach(r=>{const d=parseDateValue_(r[1]);if(!d||(start&&d<start)||(end&&d>end))return;const id=String(r[0]);if(seen[id])return;seen[id]=1;const rows=sh.getRange(2,1,last-1,CASHIER_TX_HEADERS.length).getValues().filter(x=>String(x[0])===id);const t=rows.reduce((a,x)=>a+number_(x[7]),0), paid=Math.min(t,number_(r[9]));totalPaid+=paid;totalReceivable+=Math.max(0,t-paid);});
  return {data:Object.values(map),grandTotal:grandTotal,grandProfit:0,grandModal:grandTotal,totalPaid:totalPaid,totalReceivable:totalReceivable};
}

// ==========================================================
// DATA TERBARU KASIR - KHUSUS KASIR_TRANSAKSI
// ==========================================================
function getCashierRecentData(startDate,endDate){
  const sh=getOrCreateCashierTx_(),last=sh.getLastRow();if(last<2)return [];
  const start=startDate?startOfDay_(parseDateInput(startDate)):null,end=endDate?endOfDay_(parseDateInput(endDate)):null;
  const map={};
  sh.getRange(2,1,last-1,CASHIER_TX_HEADERS.length).getValues().forEach((r,i)=>{
    const d=parseDateValue_(r[1]);if(!d)return;const day=startOfDay_(d);if(start&&day<start)return;if(end&&day>end)return;
    const id=String(r[0]||'');if(!map[id])map[id]={id:id,tanggalHari:formatDateWithDay_(d),waktu:String(r[2]||''),member:String(r[10]||''),status:normalizeStatus_(r[8]||'LUNAS'),jumlahDibayar:number_(r[9]),diskon:0,total:0,items:[],rows:[]};
    map[id].total+=number_(r[7]);map[id].diskon+=number_(r[11]);map[id].items.push(String(r[3]||'')+' '+number_(r[4])+' '+String(r[5]||''));map[id].rows.push(i+2);
  });
  return Object.values(map).sort((a,b)=>String(b.id).localeCompare(String(a.id)));
}
function tandaiLunasCashier(id){
  id=String(id||'').trim();
  if(!id) throw new Error('ID transaksi Kasir tidak valid.');
  const sh=getOrCreateCashierTx_(), last=sh.getLastRow();
  if(last<2) throw new Error('Transaksi Kasir tidak ditemukan.');
  const values=sh.getRange(2,1,last-1,CASHIER_TX_HEADERS.length).getValues();
  const rows=[]; let total=0;
  values.forEach((r,i)=>{ if(String(r[0]||'')===id){ rows.push(i+2); total+=number_(r[7]); } });
  if(!rows.length) throw new Error('Transaksi Kasir tidak ditemukan.');
  rows.forEach(row=>{
    sh.getRange(row,9).setValue('LUNAS');
    sh.getRange(row,10).setValue(total);
  });
  return {success:true,message:'Transaksi Kasir berhasil ditandai LUNAS.',recentData:getCashierRecentData('','')};
}

function deleteCashierTransaction(id){
  id=String(id||'').trim();if(!id)throw new Error('ID transaksi Kasir tidak valid.');
  const sh=getOrCreateCashierTx_(),last=sh.getLastRow();if(last<2)return {success:true,message:'Transaksi sudah tidak ada.',recentData:[]};
  const ids=sh.getRange(2,1,last-1,1).getDisplayValues().flat();
  for(let i=ids.length-1;i>=0;i--){if(String(ids[i])===id)sh.deleteRow(i+2);}
  return {success:true,message:'Transaksi Kasir berhasil dihapus.',recentData:getCashierRecentData('','')};
}

// ==========================================================
// UPLOAD FOTO PRODUK KASIR KE GOOGLE DRIVE
// Folder: 1xW3P1kwXKmac2_qm1okp80BZcsZBkTQm
// ==========================================================
function uploadCashierProductImage(payload) {
  payload = payload || {};
  const folderId = '1xW3P1kwXKmac2_qm1okp80BZcsZBkTQm';
  const name = String(payload.name || 'produk.jpg').replace(/[\\/:*?"<>|]/g, '_');
  const mimeType = String(payload.mimeType || 'image/jpeg');
  const base64 = String(payload.base64 || '');
  if (!base64) throw new Error('Data gambar tidak ditemukan.');
  if (base64.length > 8 * 1024 * 1024) throw new Error('Ukuran gambar terlalu besar. Maksimal sekitar 5 MB.');

  const bytes = Utilities.base64Decode(base64);
  const blob = Utilities.newBlob(bytes, mimeType, name);
  const folder = DriveApp.getFolderById(folderId);
  const file = folder.createFile(blob);

  // Agar gambar dapat tampil pada Web App. Pada akun Workspace tertentu,
  // administrator dapat melarang berbagi "siapa saja yang memiliki link".
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {
    // File tetap tersimpan. Jika kebijakan domain melarang public link,
    // pemilik dapat mengubah izin file secara manual di Drive.
  }

  const id = file.getId();
  return {
    success: true,
    fileId: id,
    name: file.getName(),
    url: 'https://drive.google.com/thumbnail?id=' + encodeURIComponent(id) + '&sz=w1000',
    driveUrl: file.getUrl()
  };
}


// ==========================================================
// PENGELUARAN UANG + KEUANGAN GABUNGAN PEMASUKAN/PENGELUARAN
// ==========================================================
const EXPENSE_HEADERS = ['ID','Tanggal','Keterangan','Nominal','Catatan','Waktu Input','Sumber Uang'];

function getOrCreateExpenseSheet_(){
  const ss=getSpreadsheet_();
  let sh=ss.getSheetByName(CONFIG.EXPENSE_SHEET);
  if(!sh) sh=ss.insertSheet(CONFIG.EXPENSE_SHEET);
  if(sh.getLastRow()===0){
    if(sh.getMaxColumns()<EXPENSE_HEADERS.length) sh.insertColumnsAfter(sh.getMaxColumns(),EXPENSE_HEADERS.length-sh.getMaxColumns());
    sh.getRange(1,1,1,EXPENSE_HEADERS.length).setValues([EXPENSE_HEADERS]);
  }
  sh.setFrozenRows(1);
  return sh;
}
function parseExpenseDate_(value){
  // Google Sheets biasanya mengembalikan tanggal asli sebagai Date.
  if(value instanceof Date && !isNaN(value.getTime())) return value;

  // Dukungan serial date Google Sheets/Excel jika suatu sheet menyimpan angka tanggal.
  if(typeof value === 'number' && isFinite(value) && value > 1000){
    const epoch = new Date(1899, 11, 30);
    const d = new Date(epoch.getTime() + Math.round(value) * 86400000);
    if(!isNaN(d.getTime())) return d;
  }

  const text=String(value||'').trim();
  if(!text) throw new Error('Tanggal kosong.');

  // yyyy-mm-dd / yyyy/mm/dd
  let m=text.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})(?:\s.*)?$/);
  if(m){
    const d=new Date(Number(m[1]),Number(m[2])-1,Number(m[3]));
    if(!isNaN(d.getTime())) return d;
  }

  // Mendukung dd/mm/yyyy DAN mm/dd/yyyy.
  // Jika bagian kedua > 12, pasti format mm/dd/yyyy (contoh 09/28/2026).
  // Jika bagian pertama > 12, pasti format dd/mm/yyyy.
  // Jika keduanya <= 12, aplikasi mengutamakan format Indonesia dd/mm/yyyy.
  m=text.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})(?:\s.*)?$/);
  if(m){
    const a=Number(m[1]), b=Number(m[2]), y=Number(m[3]);
    let day, month;
    if(b>12 && a<=12){ month=a; day=b; }
    else { day=a; month=b; }
    const d=new Date(y,month-1,day);
    if(d.getFullYear()===y && d.getMonth()===month-1 && d.getDate()===day) return d;
  }

  const d=new Date(text);
  if(!isNaN(d.getTime())) return d;
  throw new Error('Format tanggal pengeluaran tidak dikenali: '+text);
}
function expenseColumnMap_(sh){
  const lastCol=Math.max(1,sh.getLastColumn());
  const scanRows=Math.min(Math.max(sh.getLastRow(),1),10);
  const grid=sh.getRange(1,1,scanRows,lastCol).getDisplayValues();
  const aliases={
    id:['id','no','nomor'], tanggal:['tanggal','tgl','date'],
    keterangan:['keterangan','keperluan','keterangan / keperluan','uraian','deskripsi'],
    nominal:['nominal','jumlah','pengeluaran','jumlah pengeluaran','nominal pengeluaran','total pengeluaran','uang keluar'],
    catatan:['catatan','note','keterangan tambahan'], waktu:['waktu input','waktu','timestamp','jam'], kategori:['kategori'], sumber:['sumber uang','sumber','jenis uang']
  };
  const findIn=(h,names)=>{for(const n of names){const i=h.indexOf(n);if(i>=0)return i;}return -1;};
  for(let r=0;r<grid.length;r++){
    const h=grid[r].map(x=>String(x||'').trim().toLowerCase());
    const tanggal=findIn(h,aliases.tanggal), nominal=findIn(h,aliases.nominal), ket=findIn(h,aliases.keterangan);
    if(tanggal>=0 && nominal>=0){
      return {headerRow:r+1,headers:h,id:findIn(h,aliases.id),tanggal,keterangan:ket,nominal,catatan:findIn(h,aliases.catatan),waktu:findIn(h,aliases.waktu),kategori:findIn(h,aliases.kategori),sumber:findIn(h,aliases.sumber)};
    }
  }
  // Fallback untuk sheet lama tanpa header standar: A=ID/No, B=Tanggal, C=Keterangan, D=Nominal, E=Catatan.
  return {headerRow:1,headers:[],id:0,tanggal:1,keterangan:2,nominal:3,catatan:4,waktu:5,kategori:-1,sumber:6,fallback:true};
}
function saveExpense(data){
  data=data||{}; const tanggal=parseDateInput(data.tanggal);
  const keterangan=String(data.keterangan||'').trim(), nominal=number_(data.nominal), catatan=String(data.catatan||'').trim(), sumberUang=String(data.sumberUang||'MODAL').trim().toUpperCase()==='LABA'?'LABA':'MODAL';
  if(!keterangan) throw new Error('Keterangan/keperluan wajib diisi.');
  if(!(nominal>0)) throw new Error('Nominal pengeluaran harus lebih dari 0.');
  const sh=getOrCreateExpenseSheet_(), map=expenseColumnMap_(sh);
  // Jika sheet lama masih punya kolom Kategori, tetap dukung tanpa menghapus data lama.
  const width=Math.max(sh.getLastColumn(), EXPENSE_HEADERS.length), row=new Array(width).fill('');
  const id='OUT-'+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyyMMdd-HHmmss')+'-'+Math.floor(Math.random()*900+100);
  const put=(idx,val)=>{if(idx>=0)row[idx]=val;};
  put(map.id,id); put(map.tanggal,tanggal); put(map.keterangan,keterangan); put(map.nominal,nominal); put(map.catatan,catatan); put(map.waktu,new Date()); put(map.sumber,sumberUang);
  if(map.tanggal<0||map.keterangan<0||map.nominal<0){
    throw new Error('Header sheet PENGELUARAN tidak dikenali. Gunakan header: ID, Tanggal, Keterangan, Nominal, Catatan, Waktu Input. Data lama tidak dihapus.');
  } else sh.appendRow(row);
  try { atRecordSessionV1111_(sh.getName(),sh.getLastRow(),'PENGELUARAN',id); }
  catch(e) { console.log('Metadata sesi pengeluaran:',e); }
  // V10.89: pengeluaran langsung tercatat pada REKAP_KEUANGAN.
  try { syncRekapKeuangan_(); } catch (e) { console.log('Sync REKAP_KEUANGAN:', e); }
  return {success:true,message:'Pengeluaran berhasil disimpan.',data:getExpenseData('','')};
}
function getExpenseData(startDate,endDate){
  // V10.20: struktur PENGELUARAN dikunci sesuai sheet aktual:
  // A=ID, B=Tanggal, C=Keterangan, D=Nominal, E=Catatan, F=Waktu Input.
  const ss=getSpreadsheet_();
  const sh=ss.getSheetByName('PENGELUARAN');
  if(!sh) throw new Error('Sheet PENGELUARAN tidak ditemukan pada Spreadsheet utama.');

  const lastRow=sh.getLastRow();
  if(lastRow<2) return [];

  // getValues untuk tanggal/nominal asli + getDisplayValues sebagai fallback tampilan.
  const expenseCols=Math.min(7,sh.getMaxColumns());
  const raw=sh.getRange(2,1,lastRow-1,expenseCols).getValues();
  const disp=sh.getRange(2,1,lastRow-1,expenseCols).getDisplayValues();
  const start=startDate?startOfDay_(parseDateInput(startDate)):null;
  const end=endDate?startOfDay_(parseDateInput(endDate)):null;

  const moneyValue=(rv,dv)=>{
    if(typeof rv==='number' && isFinite(rv)) return rv;
    let t=String(dv||rv||'').trim().replace(/\s/g,'').replace(/^Rp/i,'').replace(/[^0-9,.-]/g,'');
    if(!t) return 0;
    if(t.indexOf(',')>=0){ t=t.replace(/\./g,'').replace(',','.'); }
    else { t=t.replace(/\./g,''); }
    const n=Number(t); return isFinite(n)?n:0;
  };

  const out=[];
  for(let i=0;i<raw.length;i++){
    const rr=raw[i], dd=disp[i];
    // Kolom B = tanggal. Utamakan Date asli, fallback ke teks 28/09/2026.
    let parsed=null;
    try{ parsed=parseExpenseDate_(rr[1] instanceof Date ? rr[1] : (dd[1]||rr[1])); }
    catch(e){
      // Jika filter kosong, baris dengan tanggal teks tetap jangan hilang diam-diam.
      if(start||end) continue;
    }
    if(parsed){
      const d=startOfDay_(parsed);
      if((start&&d<start)||(end&&d>end)) continue;
    }
    // Abaikan hanya baris yang benar-benar kosong.
    if(!String(dd[0]||'').trim() && !String(dd[1]||'').trim() && !String(dd[2]||'').trim() && !String(dd[3]||'').trim() && !String(dd[4]||'').trim()) continue;

    out.push({
      rowNumber:i+2,
      id:String(dd[0]||rr[0]||'').trim(),
      tanggal:parsed?formatDate_(parsed):String(dd[1]||rr[1]||'').trim(),
      dateInput:parsed?formatDateForInput_(parsed):'',
      keterangan:String(dd[2]||rr[2]||'').trim(),
      nominal:moneyValue(rr[3],dd[3]),
      catatan:String(dd[4]||rr[4]||'').trim(),
      waktuInput:String(dd[5]||rr[5]||'').trim(),
       sumberUang:String((dd[6]||rr[6]||'MODAL')).trim().toUpperCase()==='LABA'?'LABA':'MODAL'
    });
  }

  out.sort((a,b)=>{
    if(a.dateInput && b.dateInput) return a.dateInput.localeCompare(b.dateInput) || a.rowNumber-b.rowNumber;
    if(a.dateInput) return -1; if(b.dateInput) return 1; return a.rowNumber-b.rowNumber;
  });
  return out;
}
function deleteExpense(rowNumber){
  const sh=getOrCreateExpenseSheet_(), r=Number(rowNumber); if(!r||r<2||r>sh.getLastRow())throw new Error('Data pengeluaran tidak valid.');
  sh.deleteRow(r);
  try { syncRekapKeuangan_(); } catch (e) { console.log('Sync REKAP_KEUANGAN:', e); }
  return {success:true,message:'Pengeluaran berhasil dihapus.',data:getExpenseData('','')};
}
function getExpenseRecap(startDate,endDate){
  const rows=getExpenseData(startDate,endDate); let totalExpense=0; const daily={};
  rows.forEach(r=>{totalExpense+=number_(r.nominal);const key=r.dateInput;if(!daily[key])daily[key]={date:key,label:(r.tanggal||'').split(' ')[0],expense:0};daily[key].expense+=number_(r.nominal);});
  return {totalExpense,daily:Object.keys(daily).sort().map(k=>daily[k])};
}
// V10.79 - pembaca LANGSUNG sheet REKAP_KEUANGAN kolom A:F.
// A=TANGGAL, B=JENIS, C=SUMBER/LOKASI, D=KETERANGAN, E=UANG MASUK, F=UANG KELUAR.
function getRekapKeuanganDirectV1078_(startDate,endDate){
  const ss=getSpreadsheet_();
  const sh=ss.getSheetByName('REKAP_KEUANGAN');
  if(!sh || sh.getLastRow()<2) return [];
  const lastRow=sh.getLastRow();
  const raw=sh.getRange(2,1,lastRow-1,6).getValues();
  const disp=sh.getRange(2,1,lastRow-1,6).getDisplayValues();
  const awal=startDate?startOfDay_(parseDateInput(startDate)):null;
  const akhir=endDate?endOfDay_(parseDateInput(endDate)):null;
  const tz=Session.getScriptTimeZone()||'Asia/Jakarta';
  const out=[];
  raw.forEach(function(r,i){
    const d=parseDateValue_(r[0]) || parseDateValue_(disp[i][0]);
    if(awal && d && d.getTime()<awal.getTime()) return;
    if(akhir && d && d.getTime()>akhir.getTime()) return;
    const tanggal=d?Utilities.formatDate(d,tz,'dd/MM/yyyy'):String(disp[i][0]||'').trim();
    const jenis=String(disp[i][1]||'').trim();
    const sumber=String(disp[i][2]||'').trim();
    const keterangan=String(disp[i][3]||'').trim();
    const uangMasuk=number_(r[4]);
    const uangKeluar=number_(r[5]);
    if(!tanggal && !jenis && !sumber && !keterangan && !uangMasuk && !uangKeluar) return;
    out.push({tanggal:tanggal,jenis:jenis,sumber:sumber,keterangan:keterangan,uangMasuk:uangMasuk,uangKeluar:uangKeluar});
  });
  return out;
}

function getCombinedCashFlowRecap(startDate,endDate){
  // V10.81: JANGAN sinkronisasi/menulis REKAP_KEUANGAN saat membuka dashboard.
  // Dashboard hanya membaca sheet REKAP_KEUANGAN A:F yang sudah ada.
  // Ini mencegah request Keuangan Gabungan macet dan memastikan diagram tetap dirender.
  const sales=getSalesFinancialRecap('SEMUA SHEET',startDate,endDate);
  const cashier=getCashierFinancialRecap(startDate,endDate);
  const expense=getExpenseRecap(startDate,endDate);

  const totalOmzet=number_(sales.totalOmzet)+number_(cashier.totalOmzet);
  const totalPaid=number_(sales.totalPaid)+number_(cashier.totalPaid);
  const totalReceivable=number_(sales.totalReceivable)+number_(cashier.totalReceivable);
  const totalExpense=number_(expense.totalExpense);
  const netCash=totalPaid-totalExpense;

  // Laba Form Input Penjualan: baca langsung kolom laba dari transaksi.
  let formProfit=0;
  getTransactionSheetNames_().forEach(function(name){
    (getSheetData(name,startDate,endDate)||[]).forEach(function(r){
      formProfit+=number_(r.laba);
    });
  });

  // Laba Kasir: gunakan nilai recap yang tersedia. Jika versi recap
  // belum menyediakan laba, nilainya 0 dan Keuangan Gabungan tetap tampil.
  const cashierProfit=number_(cashier.totalProfit||cashier.grandProfit||cashier.profit);
  const grossProfit=formProfit+cashierProfit;
  const profitWithdrawals=number_(expense.totalExpenseProfit);
  const modalWithdrawals=number_(expense.totalExpenseModal);
  const totalProfitRemaining=grossProfit-profitWithdrawals;

  const map={};
  const ensure=function(x){
    const key=String(x.date||'');
    if(!map[key]) map[key]={date:key,label:x.label||key,formSales:0,cashier:0,expense:0};
    return map[key];
  };
  (sales.daily||[]).forEach(function(x){ensure(x).formSales+=number_(x.omzet);});
  (cashier.daily||[]).forEach(function(x){ensure(x).cashier+=number_(x.omzet);});
  (expense.daily||[]).forEach(function(x){ensure(x).expense+=number_(x.expense);});

  const daily=Object.keys(map).sort().map(function(k){
    const x=map[k];
    x.income=x.formSales+x.cashier;
    x.net=x.income-x.expense;
    return x;
  });

  return {
    totalIncome:totalPaid,
    totalExpense:totalExpense,
    netCash:netCash,
    totalOmzet:totalOmzet,
    totalPaid:totalPaid,
    totalReceivable:totalReceivable,
    totalGrossProfit:grossProfit,
    totalProfitWithdrawals:profitWithdrawals,
    totalModalWithdrawals:modalWithdrawals,
    totalProfitRemaining:totalProfitRemaining,
    daily:daily,
    // V10.85: REKAP tidak lagi dibawa oleh request ini; data sudah ditempel saat HTML dibuka.
    appVersion:'V10.85'
  };
}

// ============================================================
// V10.21 - PEMBACA LANGSUNG PENGELUARAN (TANPA DETEKSI HEADER)
// A=ID, B=Tanggal, C=Keterangan, D=Nominal, E=Catatan, F=Waktu Input
// Fungsi ini SELALU mengirim seluruh baris. Filter tanggal dilakukan di browser.
// ============================================================
function getPengeluaranLangsungV1021(){
  const ss=SpreadsheetApp.openById('1iH-U2bS3TtoxedqNfdh0Yh1GIUokaHCakNbboyPGLyI');
  const sh=ss.getSheetByName('PENGELUARAN');
  if(!sh) throw new Error('Sheet PENGELUARAN tidak ditemukan.');
  const lastRow=sh.getLastRow();
  if(lastRow<2) return {data:[],total:0,lastRow:lastRow,spreadsheetId:ss.getId(),sheetName:sh.getName()};
  const raw=sh.getRange(2,1,lastRow-1,6).getValues();
  const display=sh.getRange(2,1,lastRow-1,6).getDisplayValues();
  const data=[]; let total=0;
  for(let i=0;i<display.length;i++){
    const d=display[i], r=raw[i];
    if(!d.some(v=>String(v||'').trim())) continue;
    let dateObj=null;
    if(r[1] instanceof Date && !isNaN(r[1].getTime())) dateObj=r[1];
    else { try{ dateObj=parseExpenseDate_(d[1]||r[1]); }catch(e){} }
    let nominal=0;
    if(typeof r[3]==='number' && isFinite(r[3])) nominal=r[3];
    else {
      let t=String(d[3]||r[3]||'').replace(/Rp/gi,'').replace(/\s/g,'').trim();
      if(/^[-+]?\d{1,3}(\.\d{3})+(,\d+)?$/.test(t)) t=t.replace(/\./g,'').replace(',','.');
      else if(/^[-+]?\d{1,3}(,\d{3})+(\.\d+)?$/.test(t)) t=t.replace(/,/g,'');
      else t=t.replace(',','.');
      nominal=Number(t)||0;
    }
    total+=nominal;
    data.push({
      rowNumber:i+2,
      id:String(d[0]||r[0]||'').trim(),
      tanggal:dateObj?Utilities.formatDate(dateObj,'Asia/Jakarta','dd/MM/yyyy'):String(d[1]||r[1]||'').trim(),
      dateInput:dateObj?Utilities.formatDate(dateObj,'Asia/Jakarta','yyyy-MM-dd'):'',
      keterangan:String(d[2]||r[2]||'').trim(),
      nominal:nominal,
      catatan:String(d[4]||r[4]||'').trim(),
      waktuInput:String(d[5]||r[5]||'').trim()
    });
  }
  data.sort((a,b)=>(a.dateInput||'9999').localeCompare(b.dateInput||'9999')||a.rowNumber-b.rowNumber);
  return {data:data,total:total,lastRow:lastRow,spreadsheetId:ss.getId(),sheetName:sh.getName()};
}


// ============================================================
// V10.23 - DATA PENGELUARAN DASHBOARD (FIX FINAL)
// Membaca langsung PENGELUARAN!A2:F dan mengembalikan JSON STRING
// untuk menghindari masalah serialisasi google.script.run.
// ============================================================
function getDataPengeluaranDashboardV1023() {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sh = ss.getSheetByName('PENGELUARAN');
    if (!sh) {
      return JSON.stringify({success:false,message:'Sheet PENGELUARAN tidak ditemukan.',data:[],total:0,jumlah:0,lastRow:0});
    }
    var lastRow = sh.getLastRow();
    if (lastRow < 2) {
      return JSON.stringify({success:true,message:'Sheet PENGELUARAN ditemukan, tetapi belum ada data.',data:[],total:0,jumlah:0,lastRow:lastRow,sheetName:sh.getName()});
    }
    var rows = sh.getRange(2,1,lastRow-1,6).getDisplayValues();
    var data=[];
    var total=0;
    for (var i=0;i<rows.length;i++) {
      var r=rows[i];
      var hasData=false;
      for(var c=0;c<6;c++){ if(String(r[c]||'').trim()!==''){hasData=true;break;} }
      if(!hasData) continue;
      var nominalText=String(r[3]||'').trim();
      var nominal=parseNominalV1023_(nominalText);
      total+=nominal;
      var tanggal=String(r[1]||'').trim();
      data.push({
        rowNumber:i+2,
        id:String(r[0]||'').trim(),
        tanggal:tanggal,
        dateInput:tanggalKeIsoV1023_(tanggal),
        keterangan:String(r[2]||'').trim(),
        nominal:nominal,
        catatan:String(r[4]||'').trim(),
        waktuInput:String(r[5]||'').trim()
      });
    }
    return JSON.stringify({success:true,message:data.length+' transaksi berhasil dibaca dari Sheet PENGELUARAN.',data:data,total:total,jumlah:data.length,lastRow:lastRow,sheetName:sh.getName(),spreadsheetId:ss.getId()});
  } catch(err) {
    return JSON.stringify({success:false,message:'ERROR SERVER: '+(err && err.message ? err.message : String(err)),data:[],total:0,jumlah:0,lastRow:0});
  }
}

function parseNominalV1023_(value){
  var t=String(value==null?'':value).trim().replace(/Rp/gi,'').replace(/\s/g,'');
  if(!t) return 0;
  if(/^[-+]?\d{1,3}(\.\d{3})+(,\d+)?$/.test(t)) t=t.replace(/\./g,'').replace(',','.');
  else if(/^[-+]?\d{1,3}(,\d{3})+(\.\d+)?$/.test(t)) t=t.replace(/,/g,'');
  else t=t.replace(/[^0-9,.-]/g,'').replace(',','.');
  var n=Number(t); return isFinite(n)?n:0;
}

function tanggalKeIsoV1023_(value){
  var t=String(value||'').trim().split(' ')[0];
  var m=t.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
  if(m){
    var a=Number(m[1]),b=Number(m[2]),y=Number(m[3]),d=a,mo=b;
    if(a<=12 && b>12){mo=a;d=b;} // format US bila jelas
    return y+'-'+String(mo).padStart(2,'0')+'-'+String(d).padStart(2,'0');
  }
  m=t.match(/^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})$/);
  if(m) return m[1]+'-'+String(Number(m[2])).padStart(2,'0')+'-'+String(Number(m[3])).padStart(2,'0');
  return '';
}


// ============================================================
// V10.35 - LOGIN & ROLE ACCESS
// Sheet USERS: USERNAME | PASSWORD | ROLE | NAMA | STATUS
// Akun awal otomatis: admin/admin123 dan kasir/kasir123.
// Segera ubah password langsung pada Sheet USERS setelah login pertama.
// ============================================================
function ensureUsersSheet_() {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(CONFIG.USERS_SHEET);
  if (!sh) sh = ss.insertSheet(CONFIG.USERS_SHEET);
  const headers = ['USERNAME','PASSWORD','ROLE','NAMA','STATUS'];
  if (sh.getLastRow() === 0) sh.getRange(1,1,1,headers.length).setValues([headers]);
  else sh.getRange(1,1,1,headers.length).setValues([headers]);
  if (sh.getLastRow() < 2) {
    sh.getRange(2,1,2,5).setValues([
      ['admin','admin123','ADMIN','Administrator','AKTIF'],
      ['kasir','kasir123','KASIR','Kasir','AKTIF']
    ]);
  }
  sh.setFrozenRows(1);
  return sh;
}

function loginUser(username, password) {
  username = String(username || '').trim();
  password = String(password || '');
  if (!username || !password) return {success:false,message:'Username dan password wajib diisi.'};
  const sh = ensureUsersSheet_();
  const rows = sh.getLastRow() > 1 ? sh.getRange(2,1,sh.getLastRow()-1,5).getDisplayValues() : [];
  const found = rows.find(r => String(r[0]).trim().toLowerCase() === username.toLowerCase());
  if (!found || String(found[1]) !== password) return {success:false,message:'Username atau password salah.'};
  if (String(found[4] || 'AKTIF').trim().toUpperCase() !== 'AKTIF') return {success:false,message:'Akun ini tidak aktif.'};
  const role = String(found[2] || 'KASIR').trim().toUpperCase() === 'ADMIN' ? 'ADMIN' : 'KASIR';
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('LOGIN_' + token, JSON.stringify({username:found[0],role:role,name:found[3] || found[0]}), 21600);
  return {success:true,token:token,username:found[0],role:role,name:found[3] || found[0]};
}

function validateLogin(token) {
  if (!token) return {success:false};
  const raw = CacheService.getScriptCache().get('LOGIN_' + token);
  if (!raw) return {success:false};
  const u = JSON.parse(raw);
  return {success:true,token:token,username:u.username,role:u.role,name:u.name};
}

function logoutUser(token) {
  if (token) CacheService.getScriptCache().remove('LOGIN_' + token);
  return {success:true};
}



// ============================================================
// V10.60 - LAPORAN UANG MASUK & PENGELUARAN (PEMBACA LANGSUNG)
// Mengambil uang masuk langsung dari semua sheet transaksi penjualan
// + KASIR_TRANSAKSI, dan pengeluaran langsung dari Sheet PENGELUARAN.
// Filter menggunakan filter utama Keuangan Gabungan.
// ============================================================
function getCashFlowReportDirectV1060(startDate,endDate){
  const rows=[];
  let totalIncome=0, totalExpense=0;

  // 1. UANG MASUK dari semua sheet Form Input Penjualan.
  getTransactionSheetNames_().forEach(function(sheetName){
    (getSheetData(sheetName,startDate,endDate)||[]).forEach(function(r){
      const total=number_(r.totalPenjualan||r.total||0);
      const status=normalizeStatus_(r.status||'LUNAS');
      let paid=0;
      if(status==='LUNAS') paid=total;
      else if(status==='DP') paid=Math.min(total,Math.max(0,number_(r.jumlahDibayar)));
      if(paid<=0) return;
      totalIncome+=paid;
      rows.push({
        date:r.tanggalInput||'',
        tanggal:r.tanggal||r.tanggalHari||r.tanggalInput||'',
        jenis:'UANG MASUK',
        keterangan:'Penjualan '+sheetName+' - '+String(r.sayur||'').trim(),
        uangMasuk:paid,
        pengeluaran:0,
        sumber:sheetName
      });
    });
  });

  // 2. UANG MASUK dari KASIR_TRANSAKSI. Jumlah dibayar disimpan per ID,
  // sehingga transaksi dikelompokkan per ID agar tidak dihitung berulang.
  const shCash=getSpreadsheet_().getSheetByName(CONFIG.CASHIER_TX_SHEET);
  if(shCash && shCash.getLastRow()>=2){
    const vals=shCash.getRange(2,1,shCash.getLastRow()-1,CASHIER_TX_HEADERS.length).getValues();
    const start=startDate?startOfDay_(parseDateInput(startDate)):null;
    const end=endDate?endOfDay_(parseDateInput(endDate)):null;
    const tx={};
    vals.forEach(function(r){
      const d=parseDateValue_(r[1]); if(!d)return;
      const day=startOfDay_(d); if(start&&day<start)return; if(end&&day>end)return;
      const id=String(r[0]||'').trim()||('ROW-'+Utilities.getUuid());
      if(!tx[id]) tx[id]={id:id,date:d,total:0,paid:number_(r[9]),status:normalizeStatus_(r[8]||'LUNAS'),items:[]};
      tx[id].total+=number_(r[7]);
      if(r[3]) tx[id].items.push(String(r[3]));
      tx[id].paid=Math.max(tx[id].paid,number_(r[9]));
    });
    Object.keys(tx).forEach(function(id){
      const t=tx[id]; let paid=0;
      if(t.status==='LUNAS') paid=t.total;
      else if(t.status==='DP') paid=Math.min(t.total,Math.max(0,t.paid));
      if(paid<=0)return;
      totalIncome+=paid;
      rows.push({date:Utilities.formatDate(t.date,Session.getScriptTimeZone(),'yyyy-MM-dd'),tanggal:Utilities.formatDate(t.date,Session.getScriptTimeZone(),'dd/MM/yyyy'),jenis:'UANG MASUK',keterangan:'Kasir / Transaksi Cepat - '+id,uangMasuk:paid,pengeluaran:0,sumber:'KASIR_TRANSAKSI'});
    });
  }

  // 3. PENGELUARAN langsung dari Sheet PENGELUARAN.
  (getExpenseData(startDate,endDate)||[]).forEach(function(r){
    const nominal=number_(r.nominal); if(nominal<=0)return;
    totalExpense+=nominal;
    rows.push({date:r.dateInput||'',tanggal:r.tanggal||r.dateInput||'',jenis:'PENGELUARAN',keterangan:String(r.keterangan||'Pengeluaran')+(r.catatan?' - '+r.catatan:''),uangMasuk:0,pengeluaran:nominal,sumber:'PENGELUARAN'});
  });

  rows.sort(function(a,b){
    const da=String(a.date||''),db=String(b.date||'');
    if(da!==db)return da.localeCompare(db);
    return String(a.jenis||'').localeCompare(String(b.jenis||''));
  });
  return {rows:rows,totalIncome:totalIncome,totalExpense:totalExpense,net:totalIncome-totalExpense};
}

function getCombinedCashFlowTransactions(startDate,endDate){
  // V10.62 - sumber data laporan sesuai struktur spreadsheet aktual:
  // - KRATON / KAUMAN / RANGGEH / lokasi baru dari "Tambah Lokasi Diluar": kolom H
  // - KASIR_TRANSAKSI: kolom H
  // - PENGELUARAN: kolom D
  // Tanggal seluruh sumber dibaca dari kolom B.
  const ss=getSpreadsheet_();
  const tz=Session.getScriptTimeZone();
  const start=startDate?startOfDay_(parseDateInput(startDate)):null;
  const end=endDate?endOfDay_(parseDateInput(endDate)):null;
  const rows=[];
  let totalIncome=0,totalExpense=0;

  function inPeriod_(value){
    const d=parseDateValue_(value);
    if(!d) return null;
    const day=startOfDay_(d);
    if(start&&day<start) return null;
    if(end&&day>end) return null;
    return d;
  }
  function pushIncome_(sheetName,date,amount,ket){
    amount=number_(amount); if(!(amount>0)) return;
    totalIncome+=amount;
    rows.push({
      date:Utilities.formatDate(date,tz,'yyyy-MM-dd'),
      tanggal:Utilities.formatDate(date,tz,'dd/MM/yyyy'),
      jenis:'UANG MASUK',
      keterangan:ket||('Pemasukan '+sheetName),
      uangMasuk:amount,pengeluaran:0,sumber:sheetName
    });
  }

  // Semua sheet transaksi lokasi otomatis ikut, termasuk lokasi yang dibuat
  // melalui fitur Tambah Lokasi Diluar. Nilai Uang Masuk = KOLOM H.
  getTransactionSheetNames_().forEach(function(sheetName){
    const sh=ss.getSheetByName(sheetName); if(!sh||sh.getLastRow()<2) return;
    const width=Math.max(8,Math.min(sh.getLastColumn(),TX_HEADERS.length));
    const vals=sh.getRange(2,1,sh.getLastRow()-1,width).getValues();
    vals.forEach(function(r){
      const d=inPeriod_(r[1]); if(!d) return; // B = Tanggal
      const nominal=number_(r[7]);             // H = Uang Masuk / Total Penjualan
      const produk=String(r[2]||'').trim();
      pushIncome_(sheetName,d,nominal,'Penjualan '+sheetName+(produk?' - '+produk:''));
    });
  });

  // KASIR_TRANSAKSI: B = Tanggal, H = Subtotal/Uang Masuk.
  const cash=ss.getSheetByName(CONFIG.CASHIER_TX_SHEET);
  if(cash&&cash.getLastRow()>=2){
    const vals=cash.getRange(2,1,cash.getLastRow()-1,Math.max(8,Math.min(cash.getLastColumn(),CASHIER_TX_HEADERS.length))).getValues();
    vals.forEach(function(r){
      const d=inPeriod_(r[1]); if(!d) return;
      const nominal=number_(r[7]); // H
      const produk=String(r[3]||'').trim();
      pushIncome_('KASIR_TRANSAKSI',d,nominal,'Kasir / Transaksi Cepat'+(produk?' - '+produk:''));
    });
  }

  // PENGELUARAN: B = Tanggal, C = Keterangan, D = Nominal.
  const exp=ss.getSheetByName(CONFIG.EXPENSE_SHEET);
  if(exp&&exp.getLastRow()>=2){
    const vals=exp.getRange(2,1,exp.getLastRow()-1,Math.max(4,Math.min(exp.getLastColumn(),7))).getValues();
    vals.forEach(function(r){
      const d=inPeriod_(r[1]); if(!d) return;
      const nominal=number_(r[3]); // D
      if(!(nominal>0)) return;
      totalExpense+=nominal;
      rows.push({
        date:Utilities.formatDate(d,tz,'yyyy-MM-dd'),
        tanggal:Utilities.formatDate(d,tz,'dd/MM/yyyy'),
        jenis:'PENGELUARAN',
        keterangan:String(r[2]||'Pengeluaran').trim()||'Pengeluaran',
        uangMasuk:0,pengeluaran:nominal,sumber:'PENGELUARAN'
      });
    });
  }

  rows.sort(function(a,b){
    if(a.date!==b.date) return a.date.localeCompare(b.date);
    return String(a.jenis).localeCompare(String(b.jenis));
  });
  return {rows:rows,totalIncome:totalIncome,totalExpense:totalExpense,net:totalIncome-totalExpense};
}


// V10.65 - daftar baru: baca langsung kolom spreadsheet.
// Uang masuk: kolom H pada sheet lokasi + KASIR_TRANSAKSI.
// Uang keluar: kolom D pada PENGELUARAN.
// Filter tanggal memakai kolom B.
function getCombinedMoneyList(startDate,endDate){
  syncRekapKeuangan_();
  const ss=getSpreadsheet_();
  const sh=ss.getSheetByName('REKAP_KEUANGAN');
  if(!sh || sh.getLastRow()<2) return {rows:[],totalIncome:0,totalExpense:0};
  const vals=sh.getRange(2,1,sh.getLastRow()-1,7).getValues();
  const awal=parseFilterDate_(startDate,false), akhir=parseFilterDate_(endDate,true), rows=[];
  vals.forEach(function(v){
    const dt=parseAnyDate_(v[0]);
    if(awal&&(!dt||dt<awal))return;
    if(akhir&&(!dt||dt>akhir))return;
    rows.push({sortDate:dt?dt.getTime():0,tanggal:dt?Utilities.formatDate(dt,Session.getScriptTimeZone()||'Asia/Jakarta','dd/MM/yyyy'):String(v[0]||''),jenis:String(v[1]||''),sumber:String(v[2]||''),keterangan:String(v[3]||''),uangMasuk:number_(v[4]),uangKeluar:number_(v[5])});
  });
  rows.sort(function(a,b){return b.sortDate-a.sortDate;});
  let totalIncome=0,totalExpense=0;
  rows.forEach(function(r){totalIncome+=number_(r.uangMasuk);totalExpense+=number_(r.uangKeluar);});
  return {rows:rows,totalIncome:totalIncome,totalExpense:totalExpense};
}


function parseFilterDate_(value,endOfDay){
  if(!value)return null;
  const p=String(value).split('-');
  if(p.length===3){
    const d=new Date(Number(p[0]),Number(p[1])-1,Number(p[2]),endOfDay?23:0,endOfDay?59:0,endOfDay?59:0,endOfDay?999:0);
    return isNaN(d.getTime())?null:d;
  }
  return parseAnyDate_(value);
}
function parseAnyDate_(value){
  if(value instanceof Date && !isNaN(value.getTime()))return new Date(value.getTime());
  if(value===null||value===undefined||value==='')return null;
  const s=String(value).trim();
  let m=s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if(m){const d=new Date(+m[3],+m[2]-1,+m[1]);return isNaN(d.getTime())?null:d;}
  m=s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if(m){const d=new Date(+m[1],+m[2]-1,+m[3]);return isNaN(d.getTime())?null:d;}
  const d=new Date(s); return isNaN(d.getTime())?null:d;
}


function getOrCreateRekapKeuangan_(){
  const ss=getSpreadsheet_();
  let sh=ss.getSheetByName('REKAP_KEUANGAN');
  if(!sh) sh=ss.insertSheet('REKAP_KEUANGAN');
  const hd=['TANGGAL','JENIS','SUMBER / LOKASI','KETERANGAN','UANG MASUK','UANG KELUAR','KUNCI DATA'];
  if(sh.getMaxColumns()<7)sh.insertColumnsAfter(sh.getMaxColumns(),7-sh.getMaxColumns());
  sh.getRange(1,1,1,7).setValues([hd]); sh.setFrozenRows(1); return sh;
}
function syncRekapKeuangan_(){
  const ss=getSpreadsheet_();
  const out=[], target=getOrCreateRekapKeuangan_(), dyn={};
  try{(getTransactionSheetNames_()||[]).forEach(function(n){dyn[String(n).trim().toUpperCase()]=true;});}catch(e){}
  const known={KRATON:true,KAUMAN:true,RANGGEH:true,KASIR_TRANSAKSI:true};
  function dt(v){return parseAnyDate_(v[1])||parseAnyDate_(v[0])||parseAnyDate_(v[2]);}
  ss.getSheets().forEach(function(sh){
    const nm=sh.getName(), up=nm.trim().toUpperCase(), lr=sh.getLastRow(), lc=sh.getLastColumn();
    if(up==='REKAP_KEUANGAN'||lr<2)return;
    if(up==='PENGELUARAN'&&lc>=4){
      sh.getRange(2,1,lr-1,Math.max(4,Math.min(lc,8))).getValues().forEach(function(v,i){
        const n=number_(v[3]); if(n<=0)return; const x=dt(v);
        out.push([x||v[1],'UANG KELUAR','PENGELUARAN',String(v[2]||v[4]||'Pengeluaran'),0,n,'PENGELUARAN|'+(i+2)]);
      }); return;
    }
    if(!(known[up]||dyn[up])||lc<8)return;
    sh.getRange(2,1,lr-1,8).getValues().forEach(function(v,i){
      const n=number_(v[7]); if(n<=0)return; const x=dt(v);
      let ket='Transaksi '+nm;
      for(let j=2;j<=6;j++){const q=String(v[j]||'').trim();if(q&&!/^[\d.,\s]+$/.test(q)){ket=q;break;}}
      out.push([x||v[1],'UANG MASUK',nm,ket,n,0,up+'|'+(i+2)]);
    });
  });
  if(target.getLastRow()>1)target.getRange(2,1,target.getLastRow()-1,7).clearContent();
  if(out.length){
    if(target.getMaxRows()<out.length+1)target.insertRowsAfter(target.getMaxRows(),out.length+1-target.getMaxRows());
    target.getRange(2,1,out.length,7).setValues(out);
    target.getRange(2,1,out.length,1).setNumberFormat('dd/MM/yyyy');
    target.getRange(2,5,out.length,2).setNumberFormat('#,##0');
  }
  return {success:true,total:out.length};
}
function refreshRekapKeuangan(){return syncRekapKeuangan_();}


// V10.75: pembaca READ-ONLY khusus LIST BOX REKAP_KEUANGAN.
// Tidak sync, tidak membuat/mengubah sheet, dan tidak mengubah perhitungan Keuangan Gabungan.
function getRekapKeuanganBoxDataV1075(startDate, endDate) {
  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName('REKAP_KEUANGAN');
  if (!sh) return { success:true, rows:[], message:'Sheet REKAP_KEUANGAN tidak ditemukan.' };

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { success:true, rows:[], message:'Sheet REKAP_KEUANGAN belum berisi data.' };

  // A:F = Tanggal, Jenis, Sumber/Lokasi, Keterangan, Uang Masuk, Uang Keluar
  const raw = sh.getRange(2, 1, lastRow - 1, 6).getValues();
  const display = sh.getRange(2, 1, lastRow - 1, 6).getDisplayValues();
  const awal = parseFilterDate_(startDate, false);
  const akhir = parseFilterDate_(endDate, true);
  const tz = Session.getScriptTimeZone() || 'Asia/Jakarta';
  const rows = [];

  raw.forEach(function(r, i) {
    const d = parseAnyDate_(r[0]) || parseAnyDate_(display[i][0]);
    if (awal && (!d || d < awal)) return;
    if (akhir && (!d || d > akhir)) return;

    const masuk = number_(r[4]);
    const keluar = number_(r[5]);
    // Lewati baris yang benar-benar kosong.
    if (!d && !display[i][1] && !display[i][2] && !display[i][3] && !masuk && !keluar) return;

    rows.push({
      tanggal: d ? Utilities.formatDate(d, tz, 'dd/MM/yyyy') : String(display[i][0] || ''),
      jenis: String(display[i][1] || ''),
      sumber: String(display[i][2] || ''),
      keterangan: String(display[i][3] || ''),
      uangMasuk: masuk,
      uangKeluar: keluar,
      _sort: d ? d.getTime() : 0
    });
  });

  rows.sort(function(a,b){ return b._sort - a._sort; });
  rows.forEach(function(r){ delete r._sort; });
  return { success:true, rows:rows, message:'' };
}



// ============================================================
// V10.82 - PEMBACA SEDERHANA REKAP_KEUANGAN A:F (READ ONLY)
// Hanya membaca sheet yang sudah ada. Tidak sync / tidak menulis apa pun.
// ============================================================
function getRekapKeuanganSimple(startDate, endDate) {
  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName('REKAP_KEUANGAN');
  if (!sh) return {success:false, message:'Sheet REKAP_KEUANGAN tidak ditemukan.', rows:[], totalMasuk:0, totalKeluar:0};
  const lr = sh.getLastRow();
  if (lr < 2) return {success:true, rows:[], totalMasuk:0, totalKeluar:0};
  const vals = sh.getRange(2,1,lr-1,6).getValues();
  const disp = sh.getRange(2,1,lr-1,6).getDisplayValues();
  const tz = Session.getScriptTimeZone() || 'Asia/Jakarta';
  const awal = startDate ? new Date(startDate + 'T00:00:00') : null;
  const akhir = endDate ? new Date(endDate + 'T23:59:59') : null;
  const rows=[]; let totalMasuk=0, totalKeluar=0;
  vals.forEach(function(r,i){
    let d = r[0] instanceof Date && !isNaN(r[0].getTime()) ? r[0] : parseDateValue_(r[0]) || parseDateValue_(disp[i][0]);
    if (awal && d && d < awal) return;
    if (akhir && d && d > akhir) return;
    const masuk = number_(r[4]);
    const keluar = number_(r[5]);
    const tanggal = d ? Utilities.formatDate(d,tz,'dd/MM/yyyy') : String(disp[i][0]||'').trim();
    const jenis = String(disp[i][1]||'').trim();
    const sumber = String(disp[i][2]||'').trim();
    const keterangan = String(disp[i][3]||'').trim();
    if(!tanggal && !jenis && !sumber && !keterangan && !masuk && !keluar) return;
    totalMasuk += masuk; totalKeluar += keluar;
    rows.push({tanggal:tanggal, jenis:jenis, sumber:sumber, keterangan:keterangan, uangMasuk:masuk, uangKeluar:keluar});
  });
  return {success:true, rows:rows, totalMasuk:totalMasuk, totalKeluar:totalKeluar, appVersion:'V10.82'};
}


// ============================================================
// V10.86 - TOTAL KEUANGAN
// Dashboard baru pengganti KEUANGAN GABUNGAN.
// Menggabungkan transaksi uang masuk dari semua lokasi + kasir
// dan uang keluar dari PENGELUARAN menjadi buku kas kronologis.
// ============================================================
// V11.12: label asal transaksi hanya untuk tampilan Total Keuangan.
// Data asli dan keterangan historis pada Spreadsheet tetap utuh.
function atFinanceCategoryV1112_(r) {
  const source=String(r.sumber||'').toUpperCase();
  const origin=String(r._atSource||'').toUpperCase();
  const detail=String(r.keterangan||'').toUpperCase();
  if (source==='PENGELUARAN' || number_(r.uangKeluar)>0) return 'TRANSAKSI PENGELUARAN';
  if (source==='KASIR') return 'KASIR';
  if (source==='UANG_MASUK' || r._atFinanceManual) {
    return /TRANSFER\s*BARANG/.test(source+' '+detail) ? 'TRANSFER BARANG' : 'UANG MASUK';
  }
  if (/TRANSFER\s*BARANG/.test(origin+' '+source)) return 'TRANSFER BARANG';
  return 'FORM PENJUALAN PONDOK';
}

function getTotalKeuangan(startDate, endDate) {
  // V10.92: TOTAL KEUANGAN TANPA REKAP_KEUANGAN.
  // Uang masuk dibaca langsung dari semua sheet transaksi lokasi + KASIR_TRANSAKSI.
  // Uang keluar dibaca langsung dari PENGELUARAN.
  const ss = getSpreadsheet_();
  const tz = Session.getScriptTimeZone() || 'Asia/Jakarta';
  const awal = startDate ? startOfDay_(parseDateInput(startDate)) : null;
  const akhir = endDate ? endOfDay_(parseDateInput(endDate)) : null;
  const rows = [];
  const atSessionsV1111 = atReadSessionsV1111_();
  let totalIncome = 0, totalExpense = 0, totalOmzet = 0, totalReceivable = 0, totalProfit = 0;

  function inRange_(d) {
    if (!d) return !awal && !akhir;
    const day = startOfDay_(d);
    if (awal && day < awal) return false;
    if (akhir && day > akhir) return false;
    return true;
  }
  function dateText_(d, fallback) {
    return d ? Utilities.formatDate(d, tz, 'dd/MM/yyyy') : String(fallback || '');
  }

  // 1) PENJUALAN DARI SEMUA SHEET LOKASI (KRATON, KAUMAN, RANGGEH, lokasi tambahan, dll.)
  getTransactionSheetNames_().forEach(function(name) {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    ensureTransactionHeaders_(sh);
    const lr = sh.getLastRow();
    if (lr < 2) return;
    const data = sh.getRange(2, 1, lr - 1, TX_HEADERS.length).getValues();
    data.forEach(function(r, i) {
      if (!r[1] && !r[2]) return;
      const x = normalizeTransactionRow_(r, i + 2);
      const d = parseDateValue_(r[1]);
      if (!inRange_(d)) return;

      const omzet = number_(x.totalPenjualan);
      let masuk = 0;
      if (x.status === 'LUNAS') masuk = omzet;
      else if (x.status === 'DP') masuk = Math.min(omzet, Math.max(0, number_(x.jumlahDibayar)));
      const piutang = Math.max(0, omzet - masuk);

      totalOmzet += omzet;
      totalIncome += masuk;
      totalReceivable += piutang;
      totalProfit += number_(x.laba);

      // Tetap tampilkan transaksi penjualan walau belum ada pembayaran, agar piutang terlihat dari sumbernya.
      rows.push({
        _time: d ? d.getTime() : 0,
        _order: i + 2,
        _atSession: (atSessionsV1111[name+'|'+(i+2)]||{}).session || '',
        _atSource: (atSessionsV1111[name+'|'+(i+2)]||{}).source || name,
        tanggal: dateText_(d, r[1]),
        jenis: masuk > 0 ? 'UANG MASUK' : 'PIUTANG',
        sumber: name,
        keterangan: String(x.sayur || 'Penjualan') + (x.status ? ' - ' + x.status : ''),
        uangMasuk: masuk,
        uangKeluar: 0,
        saldo: 0
      });
    });
  });

  // 2) KASIR_TRANSAKSI - gabungkan item berdasarkan ID transaksi supaya tidak dobel per barang.
  const cash = ss.getSheetByName(CONFIG.CASHIER_TX_SHEET);
  if (cash && cash.getLastRow() >= 2) {
    const raw = cash.getRange(2, 1, cash.getLastRow() - 1, CASHIER_TX_HEADERS.length).getValues();
    const tx = {};
    raw.forEach(function(r, i) {
      const d = parseDateValue_(r[1]);
      if (!d || !inRange_(d)) return;
      const id = String(r[0] || ('ROW-' + (i + 2)));
      if (!tx[id]) tx[id] = {d:d, omzet:0, paid:number_(r[9]), status:normalizeStatus_(r[8] || 'LUNAS'), items:[], order:i+2};
      tx[id].omzet += number_(r[7]);
      const item = String(r[3] || '').trim();
      if (item && tx[id].items.indexOf(item) < 0) tx[id].items.push(item);
    });
    Object.keys(tx).forEach(function(id) {
      const t = tx[id], omzet = number_(t.omzet);
      let masuk = 0;
      if (t.status === 'LUNAS') masuk = omzet;
      else if (t.status === 'DP') masuk = Math.min(omzet, Math.max(0, number_(t.paid)));
      totalOmzet += omzet;
      totalIncome += masuk;
      totalReceivable += Math.max(0, omzet - masuk);
      rows.push({
        _time:t.d.getTime(), _order:t.order,
        _atSession:(atSessionsV1111[CONFIG.CASHIER_TX_SHEET+'|'+t.order]||{}).session||'', _atSource:'KASIR',
        tanggal:dateText_(t.d,''), jenis:masuk > 0 ? 'UANG MASUK' : 'PIUTANG',
        sumber:'KASIR',
        keterangan:(t.items.length ? t.items.join(', ') : 'Penjualan Kasir') + (t.status ? ' - ' + t.status : ''),
        uangMasuk:masuk, uangKeluar:0, saldo:0
      });
    });
  }

  // 3) PENGELUARAN - baca langsung dari sheet PENGELUARAN melalui parser yang sudah dipakai aplikasi.
  let expenses = [];
  try { expenses = getExpenseData(startDate || '', endDate || '') || []; } catch(e) { expenses = []; }
  expenses.forEach(function(x, i) {
    const keluar = number_(x.nominal);
    if (!(keluar > 0)) return;
    totalExpense += keluar;
    const d = x.dateInput ? parseDateInput(x.dateInput) : parseDateValue_(x.tanggal);
    rows.push({
      _time:d ? d.getTime() : 0, _order:1000000 + i,
      _atSession:(atSessionsV1111['PENGELUARAN|'+x.rowNumber]||{}).session||'', _atSource:'PENGELUARAN',
      tanggal:x.tanggal || dateText_(d,''), jenis:'UANG KELUAR', sumber:'PENGELUARAN',
      keterangan:String(x.keterangan || x.catatan || 'Pengeluaran'),
      uangMasuk:0, uangKeluar:keluar, saldo:0
    });
  });

  // Penerimaan manual masuk saldo, bukan omzet/laba.
  getManualIncomeData(startDate,endDate).forEach(function(x,i){
    const d=parseDateInput(x.dateInput);totalIncome+=number_(x.nominal);
    rows.push({_time:d.getTime(),_order:2000000+i,tanggal:x.tanggal,jenis:'UANG MASUK',
      _atSession:(atSessionsV1111[MANUAL_INCOME_SHEET+'|'+x.rowNumber]||{}).session||'', _atSource:x.sumber,
      sumber:x.sumber,_atFinanceManual:true,keterangan:x.keterangan,uangMasuk:number_(x.nominal),uangKeluar:0,saldo:0});
  });
  // Kronologis lama -> baru supaya saldo berjalan benar.
  rows.sort(function(a,b){ return (a._time-b._time) || (a._order-b._order); });
  // Hanya baris dengan metadata V11.11 yang digabung. Riwayat lama tetap per baris.
  const displayRows = atGroupNewSessionRowsV1111_(rows);
  let saldo = 0;
  displayRows.forEach(function(r) {
    saldo += number_(r.uangMasuk) - number_(r.uangKeluar);
    r.saldo = saldo;
    r.keterangan = atFinanceCategoryV1112_(r);
    delete r._time; delete r._order; delete r._atSession; delete r._atSource;
    delete r._atFinanceManual;
  });

  return {
    success:true,
    rows:displayRows,
    totalIncome:totalIncome,
    totalExpense:totalExpense,
    balance:totalIncome-totalExpense,
    totalOmzet:totalOmzet,
    totalReceivable:totalReceivable,
    totalProfit:totalProfit,
    appVersion:'V10.93'
  };
}

// V10.94: UANG MASUK MANUAL, bukan omzet penjualan.
const MANUAL_INCOME_SHEET='UANG_MASUK';
function manualIncomeSheet_(){
 const ss=getSpreadsheet_();let sh=ss.getSheetByName(MANUAL_INCOME_SHEET);
 if(!sh){sh=ss.insertSheet(MANUAL_INCOME_SHEET);sh.appendRow(['ID','TANGGAL','SUMBER','KETERANGAN','NOMINAL','WAKTU INPUT']);}
 return sh;
}
function saveManualIncome(p){
 p=p||{};const d=parseDateInput(p.tanggal),src=String(p.sumber||'').trim(),note=String(p.keterangan||'').trim(),n=number_(p.nominal);
 if(!d||isNaN(d.getTime())||!src||!note||!isFinite(n)||n<=0)throw new Error('Tanggal, sumber, keterangan, dan nominal positif wajib diisi.');
 const lock=LockService.getScriptLock();lock.waitLock(20000);
 try{const sh=manualIncomeSheet_();sh.appendRow(['IN-'+Utilities.getUuid(),d,src,note,n,new Date()]);try{atRecordSessionV1111_(sh.getName(),sh.getLastRow(),src);}catch(e){console.log('Metadata sesi uang masuk:',e);}}
 finally{lock.releaseLock();}
 return {success:true,message:'Uang masuk berhasil disimpan.'};
}
function getManualIncomeData(startDate,endDate){
 const sh=manualIncomeSheet_();if(sh.getLastRow()<2)return [];
 const a=startDate?startOfDay_(parseDateInput(startDate)):null,b=endDate?endOfDay_(parseDateInput(endDate)):null;
 const tz=Session.getScriptTimeZone()||'Asia/Jakarta';
 return sh.getRange(2,1,sh.getLastRow()-1,6).getValues().map((r,idx)=>{
  const d=parseDateValue_(r[1]);if(!d||(a&&d<a)||(b&&d>b))return null;
  return {rowNumber:idx+2,tanggal:Utilities.formatDate(d,tz,'dd/MM/yyyy'),dateInput:Utilities.formatDate(d,tz,'yyyy-MM-dd'),sumber:String(r[2]||''),keterangan:String(r[3]||''),nominal:number_(r[4])};
 }).filter(Boolean);
}
