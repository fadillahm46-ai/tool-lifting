// ============================================================================
// BACKEND GOOGLE APPS SCRIPT - INVENTARIS TOOLS LIFTING SE
// Simpan kode ini pada editor Apps Script Spreadsheet Anda
// ============================================================================

const SHEET_USERS = 'Users';
const SHEET_BARANG = 'MasterBarang';
const SHEET_MASUK = 'TransaksiMasuk';
const SHEET_KELUAR = 'TransaksiKeluar';

/**
 * Helper untuk mengambil Sheet atau membuat sheet & header otomatis jika belum ada
 */
function getDbSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    if (sheetName === SHEET_USERS) sheet.appendRow(['Username', 'Password', 'Name']);
    if (sheetName === SHEET_BARANG) sheet.appendRow(['Kode', 'Nama', 'Kategori', 'Satuan', 'Stok']);
    if (sheetName === SHEET_MASUK || sheetName === SHEET_KELUAR) {
      sheet.appendRow(['Timestamp', 'Kode', 'Nama', 'Jumlah', 'Keterangan', 'Bukti']);
    }
  }
  return sheet;
}

/**
 * Menangani permintaan GET dari Vercel / browser
 */
function doGet(e) {
  try {
    if (e && e.parameter && e.parameter.action) {
      const action = e.parameter.action;
      let res = {};
      if (action === 'getDashboardData') res = getDashboardData();
      else if (action === 'getBarang') res = getBarang();
      else if (action === 'getTransaksi') res = getTransaksi(e.parameter.jenis || 'Masuk');
      
      return ContentService.createTextOutput(JSON.stringify(res))
        .setMimeType(ContentService.MimeType.JSON);
    }
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Jika dibuka langsung via URL Apps Script
  return HtmlService.createHtmlOutput('<h3>Backend API Tools Lifting SE aktif dan siap digunakan.</h3>')
    .setTitle('API Tools Lifting SE');
}

/**
 * Menangani permintaan POST (API) dari Vercel
 */
function doPost(e) {
  try {
    let request = {};
    if (e && e.postData && e.postData.contents) {
      request = JSON.parse(e.postData.contents);
    }
    const action = request.action;
    let responseData = { success: false, message: 'Aksi tidak valid' };

    switch (action) {
      case 'checkLogin':
        responseData = checkLogin(request.user, request.pass);
        break;
      case 'getDashboardData':
        responseData = getDashboardData();
        break;
      case 'getBarang':
        responseData = getBarang();
        break;
      case 'addBarang':
        responseData = addBarang(request.kode, request.nama, request.kategori, request.satuan, request.stok);
        break;
      case 'deleteBarang':
        responseData = deleteBarang(request.kode);
        break;
      case 'getTransaksi':
        responseData = getTransaksi(request.jenis);
        break;
      case 'prosesTransaksi':
        responseData = prosesTransaksi(request.jenis, request.kode, request.jumlah, request.keterangan, request.bukti);
        break;
      case 'deleteTransaksi':
        responseData = deleteTransaksi(request.jenis, request.rowId);
        break;
      default:
        responseData = { success: false, message: 'Action tidak dikenal: ' + action };
    }

    return ContentService.createTextOutput(JSON.stringify(responseData))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Verifikasi Login
 */
function checkLogin(user, pass) {
  try {
    const sheet = getDbSheet(SHEET_USERS);
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      sheet.appendRow(['admin', 'admin123', 'Admin Pusat']);
      if (user === 'admin' && pass === 'admin123') {
        return { success: true, name: 'Admin Pusat' };
      }
    }

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(user).trim() && String(data[i][1]).trim() === String(pass).trim()) {
        return { success: true, name: data[i][2] || 'User' };
      }
    }
    return { success: false, message: 'Username atau password salah!' };
  } catch (e) {
    return { success: false, message: 'Terjadi kesalahan sistem: ' + e.toString() };
  }
}

/**
 * Mengambil master barang
 */
function getBarang() {
  try {
    const sheet = getDbSheet(SHEET_BARANG);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];
    
    const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    const result = [];
    
    for (let i = 0; i < data.length; i++) {
      if (!data[i][0]) continue;
      result.push({
        kode: String(data[i][0]).trim(),
        nama: String(data[i][1] || ''),
        kategori: String(data[i][2] || 'Lainnya'),
        satuan: String(data[i][3] || 'Pcs'),
        stok: Number(data[i][4]) || 0
      });
    }
    return result;
  } catch (e) {
    return [];
  }
}

/**
 * Menambah / memperbarui data barang
 */
function addBarang(kode, nama, kat, sat, stok) {
  try {
    const sheet = getDbSheet(SHEET_BARANG);
    const lastRow = sheet.getLastRow();
    let found = false;
    const numericStok = parseInt(stok, 10) || 0;
    
    if (lastRow > 1) {
      const kodeList = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < kodeList.length; i++) {
        if (String(kodeList[i][0]).trim() === String(kode).trim()) {
          sheet.getRange(i + 2, 2, 1, 4).setValues([[nama, kat, sat, numericStok]]);
          found = true;
          break;
        }
      }
    }
    
    if (!found) {
      sheet.appendRow([kode, nama, kat, sat, numericStok]);
    }
    return { success: true, message: 'Data barang berhasil disimpan.' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Menghapus data barang
 */
function deleteBarang(kode) {
  try {
    const sheet = getDbSheet(SHEET_BARANG);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: false, message: 'Barang tidak ditemukan.' };
    
    const kodeList = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < kodeList.length; i++) {
      if (String(kodeList[i][0]).trim() === String(kode).trim()) {
        sheet.deleteRow(i + 2);
        return { success: true, message: 'Data berhasil dihapus.' };
      }
    }
    return { success: false, message: 'Barang tidak ditemukan.' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Mengambil riwayat transaksi (maks 50 baris terakhir)
 */
function getTransaksi(jenis) {
  try {
    const sheetName = jenis === 'Masuk' ? SHEET_MASUK : SHEET_KELUAR;
    const sheet = getDbSheet(sheetName);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];
    
    const maxRows = 50;
    const startRow = Math.max(2, lastRow - maxRows + 1);
    const numRows = lastRow - startRow + 1;
    
    const data = sheet.getRange(startRow, 1, numRows, 6).getValues();
    const result = [];
    const tz = Session.getScriptTimeZone();
    
    for (let i = data.length - 1; i >= 0; i--) {
      if (!data[i][1]) continue;
      
      let ts = data[i][0];
      if (ts instanceof Date) {
        ts = Utilities.formatDate(ts, tz, "dd/MM/yyyy HH:mm");
      }
      result.push({
        rowId: startRow + i,
        tanggal: String(ts || ''),
        kode: String(data[i][1] || ''),
        nama: String(data[i][2] || ''),
        jumlah: Number(data[i][3]) || 0,
        keterangan: String(data[i][4] || ''),
        bukti: String(data[i][5] || '')
      });
    }
    return result;
  } catch (e) {
    return [];
  }
}

/**
 * Memproses transaksi masuk/keluar & update stok master
 */
function prosesTransaksi(jenis, kode, jumlah, ket, bukti) {
  try {
    const sheetName = jenis === 'Masuk' ? SHEET_MASUK : SHEET_KELUAR;
    const sheetTrx = getDbSheet(sheetName);
    const sheetBrg = getDbSheet(SHEET_BARANG);
    
    const dataBrg = sheetBrg.getDataRange().getValues();
    let namaBarang = "";
    let stokLama = 0;
    let rowIndex = -1;
    
    for (let i = 1; i < dataBrg.length; i++) {
      if (String(dataBrg[i][0]).trim() === String(kode).trim()) {
        namaBarang = dataBrg[i][1];
        stokLama = parseInt(dataBrg[i][4], 10) || 0;
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { success: false, message: 'Barang tidak ditemukan di database.' };
    }
    
    let stokBaru = stokLama;
    const numJumlah = parseInt(jumlah, 10) || 0;
    
    if (jenis === 'Masuk') {
      stokBaru += numJumlah;
    } else {
      if (stokLama < numJumlah) {
        return { success: false, message: 'Stok tidak mencukupi untuk dikeluarkan!' };
      }
      stokBaru -= numJumlah;
    }
    
    sheetBrg.getRange(rowIndex, 5).setValue(stokBaru);
    const timestamp = new Date();
    sheetTrx.appendRow([timestamp, kode, namaBarang, numJumlah, ket, bukti || '']);
    
    return { success: true, message: 'Transaksi sukses disimpan.' };
  } catch (e) {
    return { success: false, message: 'Gagal memproses transaksi: ' + e.toString() };
  }
}

/**
 * Menghapus transaksi & mengembalikan stok master
 */
function deleteTransaksi(jenis, rowId) {
  try {
    const sheetName = jenis === 'Masuk' ? SHEET_MASUK : SHEET_KELUAR;
    const sheetTrx = getDbSheet(sheetName);
    const sheetBrg = getDbSheet(SHEET_BARANG);
    const numRowId = parseInt(rowId, 10);
    
    const trxData = sheetTrx.getRange(numRowId, 1, 1, 6).getValues()[0];
    if (!trxData[1]) {
      return { success: false, message: 'Data transaksi tidak ditemukan.' };
    }
    
    const kode = trxData[1];
    const jumlah = parseInt(trxData[3], 10) || 0;
    
    const dataBrg = sheetBrg.getDataRange().getValues();
    let rowIndexMaster = -1;
    let stokLama = 0;
    
    for (let i = 1; i < dataBrg.length; i++) {
      if (String(dataBrg[i][0]).trim() === String(kode).trim()) {
        stokLama = parseInt(dataBrg[i][4], 10) || 0;
        rowIndexMaster = i + 1;
        break;
      }
    }
    
    if (rowIndexMaster !== -1) {
      let stokBaru = stokLama;
      if (jenis === 'Masuk') {
        stokBaru -= jumlah;
        if (stokBaru < 0) stokBaru = 0;
      } else {
        stokBaru += jumlah;
      }
      sheetBrg.getRange(rowIndexMaster, 5).setValue(stokBaru);
    }
    
    sheetTrx.deleteRow(numRowId);
    return { success: true, message: 'Transaksi berhasil dihapus dan stok telah disesuaikan.' };
  } catch (e) {
    return { success: false, message: 'Gagal menghapus transaksi: ' + e.toString() };
  }
}

/**
 * Ringkasan Dashboard & Tren 7 Hari Terakhir
 */
function getDashboardData() {
  try {
    const sheetBrg = getDbSheet(SHEET_BARANG);
    const sheetMsk = getDbSheet(SHEET_MASUK);
    const sheetKlr = getDbSheet(SHEET_KELUAR);
    
    const lastBrg = sheetBrg.getLastRow();
    const lastMsk = sheetMsk.getLastRow();
    const lastKlr = sheetKlr.getLastRow();
    
    const totalBarang = Math.max(0, lastBrg - 1);
    const totalMasuk = Math.max(0, lastMsk - 1);
    const totalKeluar = Math.max(0, lastKlr - 1);
    
    const trendIn = [0, 0, 0, 0, 0, 0, 0];
    const trendOut = [0, 0, 0, 0, 0, 0, 0];
    
    const now = new Date();
    now.setHours(0, 0, 0, 0);
    const msInDay = 1000 * 60 * 60 * 24;
    
    if (lastMsk > 1) {
      const mskDates = sheetMsk.getRange(2, 1, lastMsk - 1, 1).getValues();
      for (let i = 0; i < mskDates.length; i++) {
        if (!mskDates[i][0]) continue;
        const d = new Date(mskDates[i][0]);
        d.setHours(0, 0, 0, 0);
        const diffDays = Math.floor(Math.abs(now - d) / msInDay);
        if (diffDays >= 0 && diffDays < 7) {
          trendIn[6 - diffDays]++;
        }
      }
    }
    
    if (lastKlr > 1) {
      const klrDates = sheetKlr.getRange(2, 1, lastKlr - 1, 1).getValues();
      for (let i = 0; i < klrDates.length; i++) {
        if (!klrDates[i][0]) continue;
        const d = new Date(klrDates[i][0]);
        d.setHours(0, 0, 0, 0);
        const diffDays = Math.floor(Math.abs(now - d) / msInDay);
        if (diffDays >= 0 && diffDays < 7) {
          trendOut[6 - diffDays]++;
        }
      }
    }

    return {
      totalBarang: totalBarang,
      totalMasuk: totalMasuk,
      totalKeluar: totalKeluar,
      trendIn: trendIn,
      trendOut: trendOut
    };
  } catch (e) {
    return {
      totalBarang: 0,
      totalMasuk: 0,
      totalKeluar: 0,
      trendIn: [0, 0, 0, 0, 0, 0, 0],
      trendOut: [0, 0, 0, 0, 0, 0, 0]
    };
  }
}
