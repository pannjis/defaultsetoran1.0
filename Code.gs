const ALLOWED_EMAILS = [
  'pannjisaputra@gmail.com', // Admin / Owner
  'financeptpaj@gmail.com',  // Contoh User 1
];

const SPREADSHEET_ID = '1MXdrTIL3TNq8bJylRplIzHU7GPp-7NBMBQOxYT3wzLY';

function doGet(e) {
  // 1. Ambil Email Pengunjung
  const email = Session.getActiveUser().getEmail();

  // 2. Cek Apakah Email Ada di Daftar Izin?
  // Jika TIDAK ADA (!includes), tolak aksesnya
  if (!ALLOWED_EMAILS.includes(email)) {
    return HtmlService.createHtmlOutput(
      `<div style="font-family:sans-serif; text-align:center; margin-top:50px;">
         <h2 style="color:#d63031;">⛔ AKSES DITOLAK</h2>
         <p>Maaf, email Anda <b>(${email || 'Tidak Terdeteksi'})</b> tidak memiliki izin untuk mengakses halaman ini.</p>
         <p>Silakan hubungi Administrator (pannjisaputra@gmail.com).</p>
       </div>`
    )
    .setTitle('Akses Ditolak')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  // 3. Jika ADA di Daftar, Buka Aplikasi Seperti Biasa
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Aplikasi Keuangan Multi-Cabang')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


// --- API: GET DATA ---
function getAllData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

// Fungsi Helper: Ambil data aman & Ubah Tanggal jadi Teks
  function getSheetDataSafe(sheetName, headers, defaultData) {
    let sheet = ss.getSheetByName(sheetName);
    
    // 1. Jika Sheet tidak ditemukan, buat baru
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }

    // 2. Jika Sheet kosong, buat Header & Data Default
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
      if (defaultData) {
         const id = (sheetName === 'Transaksi') ? 'TRX-INIT' : 'INIT-' + Date.now();
         sheet.appendRow([id, ...defaultData]); 
      }
    }

    // 3. Ambil data
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    // --- PERBAIKAN UTAMA DISINI ---
    // Kita proses datanya: Jika ketemu Tanggal, ubah jadi Teks (String)
    // Ini mencegah error "null" saat data dikirim ke web
    const fixedData = data.slice(1).map(row => {
      return row.map(cell => {
        // Cek apakah cell ini berisi Tanggal
        if (Object.prototype.toString.call(cell) === '[object Date]') {
          // Ubah jadi format teks YYYY-MM-DD
          return Utilities.formatDate(cell, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        }
        return cell;
      });
    });

    return fixedData;
  }

  return {
    // Parameter: (Nama Sheet, Array Header, Array Data Default Tanpa ID)
    cabang: getSheetDataSafe(
      'Cabang', 
      ['ID', 'Nama Cabang'], 
      ['Pusat'] 
    ),
    
    katMasuk: getSheetDataSafe(
      'Kategori_Pemasukan', 
      ['ID', 'Kategori Pemasukan'], 
      ['Penjualan'] 
    ),
    
    katKeluar: getSheetDataSafe(
      'Kategori_Pengeluaran', 
      ['ID', 'Kategori Pengeluaran'], 
      ['Operasional'] 
    ),
    
    metode: getSheetDataSafe(
      'Metode_Pembayaran', 
      ['ID', 'Metode Pembayaran'], 
      ['Tunai'] 
    ),
    
    // Transaksi tidak perlu data default, cukup header
    transaksi: getSheetDataSafe(
      'Transaksi', 
      ['ID', 'Tanggal', 'Cabang', 'Tipe', 'Kategori', 'Metode', 'Nominal', 'Keterangan'], 
      null 
    )
  };
}

// --- API: SAVE TRANSACTION (BUFFER) ---
function simpanTransaksiBatch(dataTransaksi) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Transaksi');
  
  // Generate ID Unik: Timestamp + Random
  const rows = dataTransaksi.map(t => {
    const id = 'TRX-' + Date.now() + '-' + Math.floor(Math.random() * 100);
    return [
      id,
      t.tanggal,
      t.cabang,
      t.tipe,
      t.kategori,
      t.metode,
      t.nominal,
      t.keterangan
    ];
  });
  
  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }
// --- TAMBAHAN PENTING: PAKSA SIMPAN AGAR DATA TIDAK STALE ---
  SpreadsheetApp.flush();

  return "Sukses";
}

// --- API: ADD MASTER DATA ---
function tambahMasterData(jenis, nilai) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheetName = '';

  switch(jenis) {
    case 'cabang': sheetName = 'Cabang'; break;
    case 'kat_masuk': sheetName = 'Kategori_Pemasukan'; break;
    case 'kat_keluar': sheetName = 'Kategori_Pengeluaran'; break;
    case 'metode': sheetName = 'Metode_Pembayaran'; break;
  }
  
  if (sheetName) {
    const sheet = ss.getSheetByName(sheetName);
    const id = 'M-' + Date.now();
    sheet.appendRow([id, nilai]);
    
    // --- TAMBAHKAN BARIS INI (PENTING!) ---
    SpreadsheetApp.flush(); 
    // --------------------------------------

    return "Berhasil menambahkan " + nilai;
  }
  return "Gagal";
}

// --- API: HAPUS LAPORAN SATU GRUP (PERBAIKAN) ---
function hapusLaporanGrup(dateStr, branch) {
  const email = Session.getActiveUser().getEmail();
  if (email !== 'pannjisaputra@gmail.com') {
    throw new Error("AKSES DITOLAK: Anda bukan Admin.");
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Transaksi');
  const data = sheet.getDataRange().getValues();
  
  let deletedCount = 0;
  for (let i = data.length - 1; i >= 1; i--) { 
    let rowDate = data[i][1];
    let rowBranch = data[i][2]; 
    let rowDateStr = Utilities.formatDate(new Date(rowDate), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (rowDateStr === dateStr && rowBranch === branch) {
      sheet.deleteRow(i + 1); 
      deletedCount++;
    }
  }
  
  // Simpan perubahan
  SpreadsheetApp.flush();

  // --- PERBAIKAN UTAMA: Kembalikan Data Terbaru ---
  return getAllData(); 
}

// --- API: UPDATE / EDIT TRANSAKSI (RETURN DATA BARU) ---
function updateTransaksiDetail(updates, idsToDelete) {
  const email = Session.getActiveUser().getEmail();
  if (email !== 'pannjisaputra@gmail.com') {
    throw new Error("AKSES DITOLAK: Anda bukan Admin.");
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Transaksi');
  
  // 1. PROSES DELETE (Per Baris)
  if (idsToDelete && idsToDelete.length > 0) {
    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      let id = data[i][0]; 
      if (idsToDelete.includes(id)) {
        sheet.deleteRow(i + 1);
      }
    }
  }

  // 2. PROSES UPDATE (Edit Nominal/Ket)
  const refreshData = sheet.getDataRange().getValues(); 
  updates.forEach(upd => {
    for (let i = 1; i < refreshData.length; i++) {
      if (refreshData[i][0] === upd.id) {
        sheet.getRange(i + 1, 7).setValue(upd.nominal);
        sheet.getRange(i + 1, 8).setValue(upd.keterangan);
        break; 
      }
    }
  });

  // 3. PAKSA SIMPAN PERUBAHAN
  SpreadsheetApp.flush(); 

  // 4. KEMBALIKAN DATA TERBARU (KUNCI PERBAIKANNYA DISINI)
  // Kita panggil fungsi getAllData() agar website menerima data fresh
  return getAllData(); 
}

// --- FUNGSI BANTUAN: CEK EMAIL USER ---
function getUserEmail() {
  return Session.getActiveUser().getEmail();
}

// --- API: PROSES RE-INPUT (Hapus Lama & Simpan Baru) ---
function prosesReInputBatch(dataTransaksi, idsHapus) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Transaksi');
  
  // 1. HAPUS DATA LAMA (Berdasarkan ID yang sedang diedit)
  // Kita looping dari bawah ke atas agar urutan baris tidak berantakan saat dihapus
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    let currentId = data[i][0];
    
    // Cek apakah ID baris ini termasuk yang mau diedit/dihapus?
    // idsHapus dikirim dari Javascript (editingIds)
    if (idsHapus.indexOf(currentId) > -1) {
      sheet.deleteRow(i + 1);
    }
  }
  
  // 2. SIMPAN DATA BARU (Hasil Edit)
  const rows = dataTransaksi.map(t => {
    // Buat ID Baru agar fresh
    const id = 'TRX-' + Date.now() + '-' + Math.floor(Math.random() * 1000);
    return [
      id,
      t.tanggal,
      t.cabang,
      t.tipe,
      t.kategori,
      t.metode,
      t.nominal,
      t.keterangan
    ];
  });
  
  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }
  
  // 3. SIMPAN PERMANEN
  SpreadsheetApp.flush(); 
  
  return "Data berhasil di-update!";
}
