/**
 * Sistem Reservasi Extra Bed - Hotel Ivory Bandung
 * Mendukung 3 jenis: Single, Dipan, Baby Cot
 * Mendukung cek per tanggal & print laporan
 */

const SPREADSHEET_ID = '1IxmwcogGD0EUYCyE9uBOCMFkOu_aTDkob4Gfxq6rf8U';
const SHEET_RESERVATIONS = 'Reservations';
const SHEET_SETTINGS = 'Settings';

const DEFAULT_SINGLE = 10;
const DEFAULT_DIPAN = 5;
const DEFAULT_BABYCOT = 3;

function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Reservasi Extra Bed — Ivory Bandung')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Setup & Migrasi Database */
function setupDatabase() {
  const ss = getSpreadsheet();

  let settingsSheet = ss.getSheetByName(SHEET_SETTINGS);
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet(SHEET_SETTINGS);
    settingsSheet.appendRow(['key', 'value']);
  }
  
  const settingData = settingsSheet.getDataRange().getValues();
  let hasNewSettings = false;
  for (let i = 0; i < settingData.length; i++) {
    if (settingData[i][0] === 'default_single') hasNewSettings = true;
    if (settingData[i][0] === 'default_extrabed') {
      settingsSheet.getRange(i + 1, 1, 1, 2).clearContent();
    }
  }
  
  if (!hasNewSettings) {
    settingsSheet.appendRow(['default_single', DEFAULT_SINGLE]);
    settingsSheet.appendRow(['default_dipan', DEFAULT_DIPAN]);
    settingsSheet.appendRow(['default_babycot', DEFAULT_BABYCOT]);
  }
  settingsSheet.getRange('A1:B1').setFontWeight('bold').setBackground('#1a2a4a').setFontColor('#ffffff');

  let resSheet = ss.getSheetByName(SHEET_RESERVATIONS);
  const newHeaders = ['id', 'guest_name', 'room_number', 'check_in', 'check_out', 'bed_type', 'quantity', 'status', 'created_at'];
  
  if (!resSheet) {
    resSheet = ss.insertSheet(SHEET_RESERVATIONS);
    resSheet.appendRow(newHeaders);
  } else {
    const firstRow = resSheet.getRange(1, 1, 1, resSheet.getLastColumn()).getValues()[0];
    if (firstRow.includes('extrabed_count')) {
      ss.deleteSheet(resSheet);
      resSheet = ss.insertSheet(SHEET_RESERVATIONS);
      resSheet.appendRow(newHeaders);
    } else if (firstRow.length < newHeaders.length) {
      resSheet.clear();
      resSheet.appendRow(newHeaders);
    }
  }
  
  resSheet.getRange('A1:I1').setFontWeight('bold').setBackground('#1a2a4a').setFontColor('#ffffff');
  resSheet.setFrozenRows(1);
  resSheet.setColumnWidth(1, 180);
  resSheet.setColumnWidth(2, 200);
  resSheet.setColumnWidth(3, 120);
  resSheet.setColumnWidth(4, 120);
  resSheet.setColumnWidth(5, 120);
  resSheet.setColumnWidth(6, 120);
  resSheet.setColumnWidth(7, 100);
  resSheet.setColumnWidth(8, 110);
  resSheet.setColumnWidth(9, 180);
}

function getSettings() {
  setupDatabase();
  const sheet = getSpreadsheet().getSheetByName(SHEET_SETTINGS);
  const data = sheet.getDataRange().getValues();
  const settings = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) settings[data[i][0]] = parseInt(data[i][1]) || 0;
  }
  return settings;
}

function updateSettings(newSettings) {
  setupDatabase();
  const sheet = getSpreadsheet().getSheetByName(SHEET_SETTINGS);
  const data = sheet.getDataRange().getValues();
  const keys = ['default_single', 'default_dipan', 'default_babycot'];
  
  keys.forEach(key => {
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        sheet.getRange(i + 1, 2).setValue(Math.max(0, parseInt(newSettings[key])));
        found = true;
        break;
      }
    }
    if (!found) {
      sheet.appendRow([key, Math.max(0, parseInt(newSettings[key]))]);
    }
  });
  return true;
}

function adjustDefaultSetting(key, delta) {
  setupDatabase();
  const sheet = getSpreadsheet().getSheetByName(SHEET_SETTINGS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      let currentVal = parseInt(data[i][1]) || 0;
      let newVal = Math.max(0, currentVal + delta);
      sheet.getRange(i + 1, 2).setValue(newVal);
      return { success: true, key: key, newValue: newVal };
    }
  }
  let newVal = Math.max(0, delta);
  sheet.appendRow([key, newVal]);
  return { success: true, key: key, newValue: newVal };
}

function addReservation(reservation) {
  setupDatabase();
  const sheet = getSpreadsheet().getSheetByName(SHEET_RESERVATIONS);

  if (!reservation.guest_name || !reservation.room_number) {
    return { success: false, message: 'Nama tamu dan nomor kamar wajib diisi' };
  }
  if (!reservation.check_in || !reservation.check_out) {
    return { success: false, message: 'Tanggal check-in dan check-out wajib diisi' };
  }

  const checkIn = new Date(reservation.check_in);
  const checkOut = new Date(reservation.check_out);
  if (checkOut <= checkIn) {
    return { success: false, message: 'Tanggal check-out harus setelah check-in' };
  }

  const validTypes = ['single', 'dipan', 'babycot'];
  if (!validTypes.includes(reservation.bed_type)) {
    return { success: false, message: 'Tipe extra bed tidak valid' };
  }

  const id = 'RES-' + new Date().getTime();
  const row = [
    id, reservation.guest_name.trim(), reservation.room_number.trim(),
    reservation.check_in, reservation.check_out, reservation.bed_type,
    parseInt(reservation.quantity) || 1, 'active', new Date().toISOString()
  ];
  sheet.appendRow(row);
  return { success: true, id: id, message: 'Reservasi berhasil ditambahkan' };
}

function getReservations() {
  setupDatabase();
  const sheet = getSpreadsheet().getSheetByName(SHEET_RESERVATIONS);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0];
  const reservations = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    const obj = {};
    for (let j = 0; j < headers.length; j++) {
      obj[headers[j]] = data[i][j];
    }
    obj.check_in = formatDate(obj.check_in);
    obj.check_out = formatDate(obj.check_out);
    obj.quantity = parseInt(obj.quantity) || 0;
    reservations.push(obj);
  }
  reservations.sort((a, b) => new Date(b.check_in) - new Date(a.check_in));
  return reservations;
}

function cancelReservation(id) {
  setupDatabase();
  const sheet = getSpreadsheet().getSheetByName(SHEET_RESERVATIONS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.getRange(i + 1, 8).setValue('cancelled');
      return { success: true, message: 'Reservasi dibatalkan' };
    }
  }
  return { success: false, message: 'Reservasi tidak ditemukan' };
}

/** Hitung sisa extrabed berdasarkan tanggal yang dipilih */
function getStatsByDate(dateStr) {
  setupDatabase();
  const settings = getSettings();
  
  const stats = {
    single: { total: settings.default_single || 0, used: 0, remaining: 0 },
    dipan: { total: settings.default_dipan || 0, used: 0, remaining: 0 },
    babycot: { total: settings.default_babycot || 0, used: 0, remaining: 0 }
  };

  let targetDate;
  if (dateStr) {
    const parts = dateStr.split('-');
    targetDate = new Date(parts[0], parts[1] - 1, parts[2]);
  } else {
    targetDate = new Date();
  }
  targetDate.setHours(0, 0, 0, 0);

  const sheet = getSpreadsheet().getSheetByName(SHEET_RESERVATIONS);
  const data = sheet.getDataRange().getValues();
  const guestList = [];

  for (let i = 1; i < data.length; i++) {
    if (!data[i][0] || data[i][7] !== 'active') continue;

    const checkIn = new Date(data[i][3]);
    checkIn.setHours(0, 0, 0, 0);
    const checkOut = new Date(data[i][4]);
    checkOut.setHours(0, 0, 0, 0);

    if (checkIn <= targetDate && checkOut > targetDate) {
      const bedType = data[i][5];
      const count = parseInt(data[i][6]) || 0;
      
      if (stats[bedType]) {
        stats[bedType].used += count;
      }

      guestList.push({
        id: data[i][0], 
        guest_name: data[i][1], 
        room_number: data[i][2],
        check_in: formatDate(data[i][3]), 
        check_out: formatDate(data[i][4]),
        bed_type: bedType, 
        quantity: count
      });
    }
  }

  Object.keys(stats).forEach(key => {
    stats[key].remaining = stats[key].total - stats[key].used;
  });

  return {
    stats: stats,
    guestList: guestList,
    date: Utilities.formatDate(targetDate, Session.getScriptTimeZone(), 'EEEE, dd MMMM yyyy')
  };
}

function formatDate(d) {
  if (!d) return '';
  if (typeof d === 'string') {
    const parsed = new Date(d);
    if (isNaN(parsed)) return d;
    d = parsed;
  }
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  return `${yyyy}-${mm}-${dd}`;
}

function init() {
  setupDatabase();
  SpreadsheetApp.flush();
  Logger.log('Database siap & terupdate di Spreadsheet ID: ' + SPREADSHEET_ID);
}