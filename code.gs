const SS_ID = '1EG3aVWpYGoGLn9fTnNAMsTWMHu4OahtR3lyx_7eqK8Q';

function doGet() {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('Sistem Pelanggaran DPM FV')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getSS() {
  return SpreadsheetApp.openById(SS_ID);
}

// Fungsi verifikasi Login dari sheet "pengguna"
function verifikasiLogin(username, password) {
  var ss = getSS();
  var sheet = ss.getSheetByName("pengguna");
  
  if (!sheet) {
    // Jika sheet tidak ada, buat default agar tidak error (Opsional)
    sheet = ss.insertSheet("pengguna");
    sheet.appendRow(["Username", "Password", "Role"]);
    sheet.appendRow(["MKEDPM", "MKEDPMVOKASI", "admin"]);
    sheet.appendRow(["FungsionarisDPM", "DPMVOKASI", "user"]);
    return verifikasiLogin(username, password); // Rekursif sekali saja
  }
  
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == username && data[i][1] == password) {
      return {
        role: data[i][2],
        username: data[i][0]
      };
    }
  }
  return null;
}

// Fungsi mengambil daftar fungsionaris untuk dropdown
function getFungsionarisList() {
  var ss = getSS();
  var sheet = ss.getSheetByName("data_fungsionaris");
  
  if (!sheet) {
    // Jika sheet belum ada, beri data kosong atau buat (Opsional)
    return [];
  }
  
  var data = sheet.getDataRange().getValues();
  var list = [];
  
  // Ambil mulai baris ke-2 (asumsi baris 1 header)
  // Berdasarkan gambar: Kolom B (Index 1) = Nama, Kolom D (Index 3) = Jabatan
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] && data[i][3]) { // Pastikan tidak kosong
      list.push({
        nama: data[i][1],
        jabatan: data[i][3]
      });
    }
  }
  return list;
}

// Perhatikan ada tambahan parameter 'poin'
function simpanPelanggaran(nama, jabatan, jenis, poin, keterangan){
  
  // Pastikan nama sheet di Spreadsheet Anda sudah benar "pelanggaran"
  var ss = getSS();
  var sheet = ss.getSheetByName("pelanggaran");
  
  // Jika sheet belum ada buat otomatis (opsional pencegahan error)
  if(!sheet) {
    sheet = ss.insertSheet("pelanggaran");
    // Buat header
    sheet.appendRow(["ID", "Nama", "Jabatan", "Pelanggaran", "Poin", "Keterangan", "Tanggal"]);
    sheet.getRange("A1:G1").setFontWeight("bold").setBackground("#f3f4f6");
  }
  
  var id = "DPM-" + new Date().getTime(); // Format ID sedikit diubah agar profesional
  
  sheet.appendRow([
    id,
    nama,
    jabatan,
    jenis,
    Number(poin), // Disimpan sebagai angka
    keterangan,
    new Date()    // Tanggal input
  ]);
  
}

// Fungsi mengambil daftar jenis pelanggaran dari sheet "data_pelanggaran"
function getJenisPelanggaranList() {
  var ss = getSS();
  var sheet = ss.getSheetByName("data_pelanggaran");
  
  if (!sheet) return [];
  
  var data = sheet.getDataRange().getValues();
  var list = [];
  
  // Ambil mulai baris ke-2 (asumsi baris 1 header)
  // Kolom A: Kategori, Kolom B: Jenis, Kolom C: Poin
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][1]) {
      list.push({
        kategori: data[i][0],
        jenis: data[i][1],
        poin: data[i][2]
      });
    }
  }
  return list;
}

// Fungsi mengambil daftar sanksi dari sheet "data_sanksi"
function getSanksiList() {
  var ss = getSS();
  var sheet = ss.getSheetByName("data_sanksi");
  
  if (!sheet) return [];
  
  var data = sheet.getDataRange().getDisplayValues(); 
  var list = [];
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][1]) {
      list.push({
        bagian: data[i][0],
        jenis: data[i][1],
        poin: data[i][2],
        deskripsi: data[i][3],
        catatan: data[i][4]
      });
    }
  }
  return list;
}

function getData(){
  try {
    var ss = getSS();
    var sheet = ss.getSheetByName("pelanggaran");
    
    if(!sheet) return []; // Hindari error jika sheet kosong/belum ada
    
    // PERBAIKAN: Menggunakan getDisplayValues() agar tanggal dikirim sebagai teks
    var data = sheet.getDataRange().getDisplayValues();
    return data;
    
  } catch(error) {
    // Menangkap error jika ada masalah lain
    return [
      ["ID", "Nama", "Jabatan", "Pelanggaran", "Poin", "Keterangan", "Tanggal"],
      ["ERROR", "Sistem Error", error.name, error.message, "0", "Cek Apps Script", "-"]
    ];
  }
}