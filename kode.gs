/**
 * ====================================================
 * BACKEND LOGIC (Code.gs)
 * ====================================================
 * STANDAR: Standalone, LockService, Batch Operations
 */

const CONFIG = {
  SHEET_NAME: 'DaftarKlien',
  SHEET_PENGATURAN: 'Pengaturan',
  HEADERS: [
    'ID Spreadsheet', 'Nama Sekolah', 'Status', 'Nama Kop', 'Desa', 'Kecamatan', 
    'Kabupaten', 'Provinsi', 'NPSN', 'Link exec', 'Ket.', 'Tahap 1', 'Tahap 2'
  ]
};

// 1. SETUP PIN PERTAMA KALI (Jalankan sekali dari editor GAS)
function setAdminPin() {
  PropertiesService.getScriptProperties().setProperty('ADMIN_PIN', '1234');
  Logger.log("SUCCESS: PIN disimpan.");
}

// 2. ENTRY POINT (API & UI)
function doGet(e) {
  // Jika diakses dengan parameter ?id=NPSN (Untuk API Routing)
  if (e.parameter && e.parameter.id) {
    return handleRoutingAPI(e.parameter.id);
  }
  // Jika diakses normal (Menampilkan UI Admin)
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Admin Panel Klien')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 3. API ROUTING LOGIC
function handleRoutingAPI(idKlien) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) throw new Error("Database tidak ditemukan.");
    
    // Mencari NPSN di Kolom I (Kolom ke-9)
    const finder = sheet.getRange("I:I").createTextFinder(idKlien).matchEntireCell(true).findNext();
    if (!finder) throw new Error(`Klien dengan NPSN "${idKlien}" tidak terdaftar.`);
    
    const row = finder.getRow();
    // PERBAIKAN: Ambil 13 kolom agar semua data tercover
    const rowData = sheet.getRange(row, 1, 1, 13).getValues()[0]; 
    
    const status = rowData[2];   // Index 2 (Status)
    const linkExec = rowData[9]; // Index 9 (Link Exec)
    
    if (status === 'Tidak aktif') throw new Error("Akses ditolak: Status klien Dinonaktifkan.");
    if (!linkExec) throw new Error("Link eksekusi belum di-input oleh Admin.");
    
    return ContentService.createTextOutput(JSON.stringify({ success: true, url: linkExec }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// 4. INCLUDE HTML PARTIALS
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// 5. AUTENTIKASI
function verifyPin(inputPin) {
  try {
    if (!inputPin) throw new Error("PIN kosong!");
    const storedPin = PropertiesService.getScriptProperties().getProperty('ADMIN_PIN');
    if (!storedPin) throw new Error("Sistem belum dikonfigurasi. Hubungi Developer.");
    
    return { success: true, isValid: inputPin === storedPin };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// 6. FETCH DATA KLIEN
function getClients() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) throw new Error(`Sheet '${CONFIG.SHEET_NAME}' tidak ditemukan!`);
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, data: [] }; // Hanya header atau kosong
    
    const clients = data.slice(1).map(row => ({
      id: String(row[0] || ''),
      namaSekolah: String(row[1] || ''),
      status: String(row[2] || ''),
      namaKop: String(row[3] || ''),
      desa: String(row[4] || ''),
      kecamatan: String(row[5] || ''),
      kabupaten: String(row[6] || ''),
      provinsi: String(row[7] || ''),
      npsn: String(row[8] || ''),
      linkExec: String(row[9] || ''),
      ket: String(row[10] || ''),
      tahap1: String(row[11] || 'Belum'),
      tahap2: String(row[12] || 'Belum')
    }));
    
    return { success: true, data: clients };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// 7. FETCH SETTINGS GLOBAL
function getGlobalSettings() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_PENGATURAN);
    if (!sheet) return { success: true, data: [] }; 
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, data: [] };
    
    const settings = data.slice(1).filter(row => row[0] && row[1]).map(row => ({
      name: String(row[0]),
      url: String(row[1])
    }));
    
    return { success: true, data: settings };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// ==========================================
// 8. DATA FORM MASUK (PENDING DATA)
// ==========================================
function getPendingFormData() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DataForm');
    if (!sheet) throw new Error("Sheet 'DataForm' tidak ditemukan. Pastikan Form tertaut ke sini.");
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, data: [] };
    
    // Ambil data NPSN dari sheet utama untuk cek duplikat secara Batch
    const mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
    let existingNPSNs = [];
    if (mainSheet) {
      const mainData = mainSheet.getDataRange().getValues();
      if (mainData.length > 1) {
        // Di DaftarKlien (tanpa Timestamp), NPSN ada di Kolom I (Index 8)
        existingNPSNs = mainData.slice(1).map(row => String(row[8] || '').trim()); 
      }
    }

    const pendingData = [];
    
    // Looping dari baris ke-2 (index 1)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      // Di DataForm (dengan Timestamp), Status Sinkronisasi ada di Kolom L (Index 11)
      const statusSinkronisasi = String(row[11] || '').trim(); 
      
      // Filter: Hanya ambil data yang belum diproses
      if (statusSinkronisasi === '') {
        // Di DataForm, NPSN ada di Kolom J (Index 9)
        const npsnVal = String(row[9] || '').trim(); 
        
        // Logika Validasi Duplikat
        const validationStatus = existingNPSNs.includes(npsnVal) ? 'Duplikat' : 'Aman';
        
        pendingData.push({
          rowIdx: i + 1, // Baris asli di spreadsheet (1-based index)
          timestamp: row[0] instanceof Date ? Utilities.formatDate(row[0], Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm") : String(row[0] || '-'),
          id: String(row[1] || ''),
          namaSekolah: String(row[2] || ''),
          status: String(row[3] || 'Aktif'),
          namaKop: String(row[4] || ''),
          desa: String(row[5] || ''),
          kecamatan: String(row[6] || ''),
          kabupaten: String(row[7] || ''),
          provinsi: String(row[8] || ''),
          npsn: npsnVal,
          linkExec: String(row[10] || ''),
          validationStatus: validationStatus
        });
      }
    }
    
    return { success: true, data: pendingData };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function markFormProcessed(rowIdx, message) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DataForm');
    if (!sheet) throw new Error("Sheet 'DataForm' tidak ditemukan.");
    
    // Tulis pesan di Kolom L (Kolom ke-12)
    sheet.getRange(rowIdx, 12).setValue(message);
    return { success: true };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// ==========================================
// 9. CRUD OPERATIONS (Dengan LockService)
// ==========================================
function addClient(clientData) {
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(10000)) throw new Error("Server sibuk. Coba lagi.");
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
    const exists = sheet.getRange("A:A").createTextFinder(clientData.id).matchEntireCell(true).findNext();
    if (exists) throw new Error("ID Spreadsheet sudah ada di database!");
    
    const rowData = [
      clientData.id, clientData.namaSekolah, clientData.status, clientData.namaKop, 
      clientData.desa, clientData.kecamatan, clientData.kabupaten, clientData.provinsi, 
      clientData.npsn, clientData.linkExec, clientData.ket, clientData.tahap1, clientData.tahap2
    ];
    
    sheet.appendRow(rowData);
    
    // LOGIKA BARU: Jika ini adalah data yang ditarik dari Form, tandai otomatis di sheet DataForm!
    if (clientData.formRowIdx) {
      const formSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DataForm');
      if (formSheet) {
        formSheet.getRange(clientData.formRowIdx, 12).setValue('Ditarik - ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy"));
      }
    }

    return { success: true, message: "Data berhasil ditambahkan!" };
  } catch (error) {
    return { success: false, message: error.message };
  } finally {
    lock.releaseLock();
  }
}

function updateClient(oldId, clientData) {
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(10000)) throw new Error("Server sibuk. Coba lagi.");
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
    const finder = sheet.getRange("A:A").createTextFinder(oldId).matchEntireCell(true).findNext();
    if (!finder) throw new Error("Data klien tidak ditemukan.");
    
    const rowIdx = finder.getRow();
    const rowData = [[
      clientData.id, clientData.namaSekolah, clientData.status, clientData.namaKop, 
      clientData.desa, clientData.kecamatan, clientData.kabupaten, clientData.provinsi, 
      clientData.npsn, clientData.linkExec, clientData.ket, clientData.tahap1, clientData.tahap2
    ]];
    
    sheet.getRange(rowIdx, 1, 1, 13).setValues(rowData);
    return { success: true, message: "Data berhasil diperbarui!" };
  } catch (error) {
    return { success: false, message: error.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteClient(id) {
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(10000)) throw new Error("Server sibuk. Coba lagi.");
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
    const finder = sheet.getRange("A:A").createTextFinder(id).matchEntireCell(true).findNext();
    if (!finder) throw new Error("Data tidak ditemukan.");
    
    sheet.deleteRow(finder.getRow());
    return { success: true, message: "Data berhasil dihapus!" };
  } catch (error) {
    return { success: false, message: error.message };
  } finally {
    lock.releaseLock();
  }
}
