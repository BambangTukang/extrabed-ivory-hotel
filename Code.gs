// ============================================
// KONFIGURASI - Ganti dengan Spreadsheet ID kamu
// ============================================
var SPREADSHEET_ID = '1caOYZOBN1UzBqcv5r_KtxCGp0qPY030wxTRFzXMws4Q';

// ============================================
// HELPER - Ambil sheet, buat jika belum ada
// ============================================
function getSheet(name, headers) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers && headers.length > 0) {
      sheet.appendRow(headers);
      var range = sheet.getRange(1, 1, 1, headers.length);
      range.setFontWeight('bold');
      range.setBackground('#1c1f2e');
      range.setFontColor('#e8e6f0');
    }
  }
  return sheet;
}

// Init semua sheet & data default saat pertama kali
function initAll() {
  getSheet('Data', ['ID', 'NamaTamu', 'NoKamar', 'CheckIn', 'CheckOut']);
  var shSet = getSheet('Settings', ['Key', 'Value']);
  var setVals = shSet.getDataRange().getValues();
  var hasDefBed = false;
  for (var i = 1; i < setVals.length; i++) {
    if (setVals[i][0] === 'defBed') { hasDefBed = true; break; }
  }
  if (!hasDefBed) shSet.appendRow(['defBed', '10']);

  var shUsers = getSheet('Users', ['Username', 'Password', 'Role']);
  var userVals = shUsers.getDataRange().getValues();
  var hasAdmin = false;
  for (var j = 1; j < userVals.length; j++) {
    if (userVals[j][0] === 'admin') { hasAdmin = true; break; }
  }
  if (!hasAdmin) shUsers.appendRow(['admin', 'admin123', 'supervisor']);
}

// ============================================
// WEB APP ENTRY (hanya doGet, doPost dihapus)
// ============================================
function doGet() {
  initAll();
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('ExtraBed Manager — Hotel')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

// ============================================
// FUNGSI YANG DIPANGGIL OLEH google.script.run
// ============================================

function gsLogin(d) {
  initAll();
  var sheet = getSheet('Users');
  var vals = sheet.getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    if (vals[i][0].toLowerCase() === d.username.toLowerCase() && vals[i][1] === d.password) {
      return { success: true, user: { username: vals[i][0], role: vals[i][2] } };
    }
  }
  return { success: false, message: 'Username atau password salah' };
}

function gsReadData() {
  initAll();
  var sheet = getSheet('Data');
  var vals = sheet.getDataRange().getValues();
  if (vals.length <= 1) return { success: true, data: [] };
  var out = [];
  for (var i = 1; i < vals.length; i++) {
    out.push({
      _row: i + 1,
      id: vals[i][0],
      name: vals[i][1],
      room: vals[i][2],
      ci: vals[i][3],
      co: vals[i][4]
    });
  }
  return { success: true, data: out };
}

function gsCreateData(d) {
  initAll();
  var sheet = getSheet('Data');
  var id = 'eb' + Date.now() + Math.random().toString(36).substr(2, 6);
  sheet.appendRow([id, d.name, d.room || '', d.ci, d.co]);
  return { success: true, message: 'Data berhasil ditambahkan', id: id };
}

function gsUpdateData(d) {
  initAll();
  var sheet = getSheet('Data');
  var row = d._row;
  if (!row || row < 2) return { success: false, message: 'Row tidak valid' };
  sheet.getRange(row, 2, 1, 4).setValues([[d.name, d.room || '', d.ci, d.co]]);
  return { success: true, message: 'Data berhasil diperbarui' };
}

function gsDeleteData(d) {
  initAll();
  var sheet = getSheet('Data');
  var row = d._row;
  if (!row || row < 2) return { success: false, message: 'Row tidak valid' };
  sheet.deleteRow(row);
  return { success: true, message: 'Data berhasil dihapus' };
}

function gsGetSetting() {
  initAll();
  var sheet = getSheet('Settings');
  var vals = sheet.getDataRange().getValues();
  var result = {};
  for (var i = 1; i < vals.length; i++) {
    result[vals[i][0]] = vals[i][1];
  }
  return { success: true, settings: result };
}

function gsSaveSetting(d) {
  initAll();
  var sheet = getSheet('Settings');
  var vals = sheet.getDataRange().getValues();
  var found = false;
  for (var i = 1; i < vals.length; i++) {
    if (vals[i][0] === d.key) {
      sheet.getRange(i + 1, 2).setValue(d.value);
      found = true;
      break;
    }
  }
  if (!found) sheet.appendRow([d.key, d.value]);
  return { success: true, message: 'Pengaturan disimpan' };
}

function gsReadUsers() {
  initAll();
  var sheet = getSheet('Users');
  var vals = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < vals.length; i++) {
    out.push({
      _row: i + 1,
      username: vals[i][0],
      password: vals[i][1],
      role: vals[i][2]
    });
  }
  return { success: true, users: out };
}

function gsAddUser(d) {
  initAll();
  var sheet = getSheet('Users');
  var vals = sheet.getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    if (vals[i][0].toLowerCase() === d.username.toLowerCase()) {
      return { success: false, message: 'Username sudah digunakan' };
    }
  }
  sheet.appendRow([d.username, d.password, d.role]);
  return { success: true, message: 'User berhasil ditambahkan' };
}

function gsDeleteUser(d) {
  initAll();
  var sheet = getSheet('Users');
  var row = d._row;
  if (!row || row < 2) return { success: false, message: 'Row tidak valid' };
  sheet.deleteRow(row);
  return { success: true, message: 'User berhasil dihapus' };
}