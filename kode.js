function doGet() {
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle('Smart RT')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 1. FUNGSI UNTUK SETUP DATABASE AWAL (WAJIB DIJALANKAN SEKALI)
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Struktur tabel beserta kolomnya
  const sheets = {
    'Users': ['id', 'username', 'password', 'role', 'nama'],
    'Warga': ['id', 'no_kk', 'nik', 'nama', 'jk', 'kerja', 'hp', 'alamat'],
    'Pengumuman': ['id', 'tgl', 'judul', 'konten', 'penulis'],
    'Keluhan': ['id', 'tgl', 'user', 'nama', 'kategori', 'desc', 'status', 'tanggapan'],
    'Surat': ['id', 'tgl', 'user', 'nama', 'jenis', 'ket', 'status'],
    'Keuangan': ['id', 'tgl', 'jenis', 'kategori', 'ket', 'masuk', 'keluar'],
    'Iuran': ['id', 'user', 'nama', 'bulan', 'nominal', 'status']
  };

  // Buat sheet jika belum ada
  for (let sheetName in sheets) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(sheets[sheetName]);
    }
  }

  // Buat akun Admin default jika sheet Users masih kosong (baru 1 baris header)
  let userSheet = ss.getSheetByName('Users');
  if (userSheet.getLastRow() === 1) {
    userSheet.appendRow([Date.now(), 'admin', 'edudigital', 'admin', 'Bpk. Budi (Ketua RT)']);
    userSheet.appendRow([Date.now()+1, 'peserta', 'edudigital', 'warga', 'Sdr. Andi (Warga)']);
  }
}

// ============================================
// FUNGSI LOGIN & REGISTRASI (DARI FRONTEND)
// ============================================

// Fungsi untuk validasi saat tombol Masuk diklik
function checkLogin(username, password) {
  try {
    let users = getSheetDataAsObjects('Users');
    let user = users.find(u => String(u.username) === String(username) && String(u.password) === String(password));
    
    if (user) {
      return {
        status: 'success',
        nama: user.nama,
        role: user.role
      };
    } else {
      return { status: 'error', message: 'Username atau Password salah!' };
    }
  } catch (error) {
    return { status: 'error', message: error.message };
  }
}

// Fungsi Pendaftaran akun Warga baru
function registerUser(nama, nik, password) {
  try {
    appendToSheet('Users', [Date.now(), nik, password, 'warga', nama]);
    return { status: 'success' };
  } catch (error) {
    return { status: 'error', message: error.message };
  }
}

// ============================================
// 2. FUNGSI UTAMA PENERIMA DATA DARI FRONTEND (doPost)
// ============================================
function doPost(e) {
  try {
    // Tangkap request dari Frontend
    let request = JSON.parse(e.postData.contents);
    let action = request.action;
    let p = request.payload;
    let data = null;

    // Routing Logika Backend
    if (action === 'login') {
      let users = getSheetDataAsObjects('Users');
      let user = users.find(u => String(u.username) === String(p.username) && String(u.password) === String(p.password));
      if (user) {
        data = { username: user.username, role: user.role, nama: user.nama };
      } else {
        throw new Error("Username atau password salah!");
      }
    }
    // Modul Pendaftaran Akun Warga Baru
    else if (action === 'register') {
      let users = getSheetDataAsObjects('Users');
      let exist = users.find(u => String(u.username) === String(p.username));
      if (exist) {
        throw new Error("NIK / Username tersebut sudah terdaftar!");
      }
      // Simpan user dengan default role 'warga'
      appendToSheet('Users', [Date.now(), p.username, p.password, 'warga', p.nama]);
      data = true;
    }
    // Modul Warga
    else if (action === 'getWarga') { data = getSheetDataAsObjects('Warga'); }
    else if (action === 'addWarga') { appendToSheet('Warga', [Date.now(), p.no_kk, p.nik, p.nama, p.jk, p.kerja, p.hp, p.alamat]); data = true; }
    else if (action === 'delWarga') { deleteFromSheet('Warga', p.id); data = true; }
    
    // Modul Pengumuman
    else if (action === 'getPengumuman') { data = getSheetDataAsObjects('Pengumuman').reverse(); } // Reverse agar yang terbaru di atas
    else if (action === 'addBerita') { appendToSheet('Pengumuman', [Date.now(), getTgl(), p.judul, p.konten, p.penulis]); data = true; }
    
    // Modul Keluhan
    else if (action === 'getKeluhan') { data = getSheetDataAsObjects('Keluhan'); }
    else if (action === 'addKeluhan') { appendToSheet('Keluhan', [Date.now(), getTgl(), p.user, p.nama, p.kategori, p.desc, 'Menunggu', '-']); data = true; }
    else if (action === 'updateKeluhan') { updateSheetRow('Keluhan', p.id, { 'status': p.status, 'tanggapan': p.tanggapan }); data = true; }
    
    // Modul Surat
    else if (action === 'getSurat') { data = getSheetDataAsObjects('Surat'); }
    else if (action === 'addSurat') { appendToSheet('Surat', [Date.now(), getTgl(), p.user, p.nama, p.jenis, p.ket, 'Menunggu']); data = true; }
    else if (action === 'updateSurat') { updateSheetRow('Surat', p.id, { 'status': p.status }); data = true; }
    
    // Modul Keuangan
    else if (action === 'getKeuangan') { data = getSheetDataAsObjects('Keuangan'); }
    else if (action === 'addTrx') { appendToSheet('Keuangan', [Date.now(), getTgl(), p.jenis, p.kategori, p.ket, p.masuk, p.keluar]); data = true; }
    
    // Modul Iuran
    else if (action === 'getIuran') { data = getSheetDataAsObjects('Iuran'); }
    else if (action === 'updateIuran') { updateSheetRow('Iuran', p.id, { 'status': 'Lunas' }); data = true; }

    // Kembalikan Response ke Frontend
    return createResponse('success', data);
    
  } catch (error) {
    return createResponse('error', null, error.message);
  }
}

// ============================================
// HELPER FUNCTIONS (Fungsi Bantuan Database)
// ============================================

function createResponse(status, data, message = "") {
  return ContentService.createTextOutput(JSON.stringify({
    status: status,
    data: data,
    message: message
  })).setMimeType(ContentService.MimeType.JSON);
}

function getTgl() {
  return Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy");
}

function getSheetDataAsObjects(sheetName) {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return [];
  let data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  
  let headers = data[0];
  let result = [];
  for (let i = 1; i < data.length; i++) {
    let obj = {};
    for (let j = 0; j < headers.length; j++) {
      obj[headers[j]] = data[i][j];
    }
    result.push(obj);
  }
  return result;
}

function appendToSheet(sheetName, rowData) {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  sheet.appendRow(rowData);
}

function updateSheetRow(sheetName, id, updateObj) {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  let data = sheet.getDataRange().getValues();
  let headers = data[0];

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) { 
      for (let key in updateObj) {
        let colIndex = headers.indexOf(key);
        if (colIndex > -1) {
          sheet.getRange(i + 1, colIndex + 1).setValue(updateObj[key]);
        }
      }
      break;
    }
  }
}

function deleteFromSheet(sheetName, id) {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  let data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
}