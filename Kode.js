// ============================================================
// BACKEND: BUAT JOKE! - PENGHIBUR KREATIF AI
// Lisensi: Arraihan Academy (SWK)
// Arsitektur: GAS-PRO-API (Pure REST API JSON)
// ============================================================

const CONFIG = {
  get SPREADSHEET_ID() { return PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'); },
  get ROOT_FOLDER_ID() { return PropertiesService.getScriptProperties().getProperty('ROOT_FOLDER_ID'); },
  SHEET_NAMES: {
    config: 'AppConfig',
    data: 'DataJoke'
  }
};

// ============================================================
// ENTRY POINTS — doGet & doPost (Pure JSON REST API)
// ============================================================
function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : 'getData';
    let result;
    if (action === 'getData') {
      result = getData();
    } else if (action === 'getStats') {
      result = getStats();
    } else {
      result = { success: false, message: 'Action GET tidak dikenal: ' + action };
    }
    return buildResponse(result);
  } catch (err) {
    return buildResponse({ success: false, message: err.message });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action || 'saveData';
    let result;

    if (action === 'saveData') {
      result = saveData(payload.data);
    } else if (action === 'likeItem') {
      result = likeItem(payload.data);
    } else if (action === 'deleteItem') {
      result = deleteItem(payload.data);
    } else {
      result = { success: false, message: 'Action POST tidak dikenal: ' + action };
    }
    return buildResponse(result);
  } catch (err) {
    return buildResponse({ success: false, message: err.message });
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// BUSINESS LOGIC (CRUD & CACHE)
// ============================================================
function getData() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('joke_data');
  if (cached) {
    return { success: true, data: JSON.parse(cached) };
  }

  const sheet = getOrCreateSheet(CONFIG.SHEET_NAMES.data, [
    'id', 'timestamp', 'nama_pengguna', 'konteks_topik', 'kategori_format',
    'gaya_bahasa', 'hasil_joke', 'hasil_pantun', 'hasil_syair', 'likes'
  ]);

  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { success: true, data: [] };

  const headers = rows[0];
  const data = rows.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  }).reverse(); // Terbaru di atas

  cache.put('joke_data', JSON.stringify(data), 1800);
  return { success: true, data: data };
}

function saveData(data) {
  if (!data || !data.konteks_topik) {
    return { success: false, message: 'Topik/konteks hiburan wajib diisi!' };
  }

  const sheet = getOrCreateSheet(CONFIG.SHEET_NAMES.data, [
    'id', 'timestamp', 'nama_pengguna', 'konteks_topik', 'kategori_format',
    'gaya_bahasa', 'hasil_joke', 'hasil_pantun', 'hasil_syair', 'likes'
  ]);

  const newRow = [
    Utilities.getUuid(),
    new Date().toISOString(),
    data.nama_pengguna || 'Sahabat Arraihan',
    data.konteks_topik || '-',
    data.kategori_format || 'Paket Komplit',
    data.gaya_bahasa || 'Santai',
    data.hasil_joke || '-',
    data.hasil_pantun || '-',
    data.hasil_syair || '-',
    Number(data.likes || 1)
  ];

  sheet.appendRow(newRow);
  CacheService.getScriptCache().remove('joke_data');

  return { success: true, message: 'Hiburan berhasil disimpan ke Galeri Cloud!' };
}

function likeItem(data) {
  const sheet = getOrCreateSheet(CONFIG.SHEET_NAMES.data);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.id)) {
      const currentLikes = Number(rows[i][9] || 0) + 1;
      sheet.getRange(i + 1, 10).setValue(currentLikes);
      CacheService.getScriptCache().remove('joke_data');
      return { success: true, likes: currentLikes, message: 'Apresiasi terkirim!' };
    }
  }
  return { success: false, message: 'Data tidak ditemukan.' };
}

function deleteItem(data) {
  if (String(data.pin) !== '123456') {
    return { success: false, message: 'PIN Admin tidak valid!' };
  }
  const sheet = getOrCreateSheet(CONFIG.SHEET_NAMES.data);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.id)) {
      sheet.deleteRow(i + 1);
      CacheService.getScriptCache().remove('joke_data');
      return { success: true, message: 'Data berhasil dihapus!' };
    }
  }
  return { success: false, message: 'ID tidak ditemukan.' };
}

// ============================================================
// HELPER & AUTO-SETUP ENVIRONMENT (JALANKAN SEKALI)
// ============================================================
function buildResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet(name, headers) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers && headers.length > 0) {
      sheet.appendRow(headers);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground('#d97706')
        .setFontColor('#ffffff');
    }
  }
  return sheet;
}

function setupAppEnvironment() {
  const APP_NAME = 'Buat Joke - Arraihan Academy SWK';
  Logger.log('Memulai setup ' + APP_NAME + '...');

  const rootFolder = DriveApp.createFolder('📁 ' + APP_NAME);
  const ss = SpreadsheetApp.create('🗃️ Database — ' + APP_NAME);
  DriveApp.getFileById(ss.getId()).moveTo(rootFolder);

  const configSheet = ss.getActiveSheet().setName(CONFIG.SHEET_NAMES.config);
  configSheet.appendRow(['key', 'value', 'keterangan']);
  configSheet.appendRow(['APP_NAME', APP_NAME, 'Nama Aplikasi']);
  configSheet.appendRow(['SPREADSHEET_ID', ss.getId(), 'ID Database Spreadsheet']);
  configSheet.appendRow(['ROOT_FOLDER_ID', rootFolder.getId(), 'ID Folder Root Drive']);
  configSheet.appendRow(['ADMIN_PIN', '123456', 'PIN Admin Default']);
  configSheet.appendRow(['LICENSE_OWNER', 'Arraihan Academy (SWK)', 'Pemilik Lisensi']);
  configSheet.setFrozenRows(1);

  const dataSheet = ss.insertSheet(CONFIG.SHEET_NAMES.data);
  const headers = [
    'id', 'timestamp', 'nama_pengguna', 'konteks_topik', 'kategori_format',
    'gaya_bahasa', 'hasil_joke', 'hasil_pantun', 'hasil_syair', 'likes'
  ];
  dataSheet.appendRow(headers);
  dataSheet.setFrozenRows(1);
  dataSheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#d97706')
    .setFontColor('#ffffff');

  // Dummy Data Awal
  dataSheet.appendRow([
    Utilities.getUuid(),
    new Date().toISOString(),
    'Admin SWK',
    'Pembuka rapat koordinasi pagi',
    'Paket Komplit',
    'Kantor/Formal Santun',
    'Kenapa rapat pagi selalu butuh kopi hangat? Karena kalau cuma butuh kepastian, itu namanya nunggu honor turun sambil senyum-senyum sendiri!',
    'Pergi ke pasar membeli duku,\nJangan lupa membeli mangga.\nSelamat pagi Bapak dan Ibu,\nSemoga rapat hari ini penuh berkah dan bahagia.',
    'Pagi cerah menyambut langkah,\nDuduk bersama menyusun arah,\nKerja ikhlas membawa berkah,\nSenyum merekah hilangkan lelah.',
    12
  ]);

  PropertiesService.getScriptProperties().setProperties({
    SPREADSHEET_ID: ss.getId(),
    ROOT_FOLDER_ID: rootFolder.getId()
  });

  Logger.log('✅ Setup selesai! Spreadsheet URL: ' + ss.getUrl());
}