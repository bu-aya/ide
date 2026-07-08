/**
 * BACKEND API SISTEM ADMINISTRASI SEKOLAH (SDIT ANNIHAYAH)
 * Untuk dihubungkan dengan UI di Blogger
 */

// 1. FUNGSI SETUP DATABASE OTOMATIS
// Jalankan fungsi ini SEKALI SAJA di editor Apps Script untuk membuat Sheet beserta kolomnya
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const sheetsDef = {
    'Settings': ['ID', 'Kepsek', 'TahunAktif'],
    'Kelas': ['id', 'nama', 'wali'],
    'CP_TP': ['tingkat', 'mapel', 'cp', 'tp'],
    'Siswa': ['id', 'nis', 'nisn', 'nama', 'jk', 'tmpt_lahir', 'tgl_lahir', 'ayah', 'ibu', 'ta_masuk', 'kelas', 'status', 'no_ijazah', 'tgl_lulus', 'rata_rata', 'tgl_masuk', 'sekolah_asal', 'alamat_asal', 'tgl_keluar', 'sekolah_tujuan', 'alasan_keluar'],
    'SuratMasuk': ['id', 'no_agenda', 'no_surat', 'tgl', 'pengirim', 'perihal'],
    'SuratKeluar': ['id', 'no_urut', 'tgl', 'jenis', 'tujuan']
  };

  for (const [sheetName, headers] of Object.entries(sheetsDef)) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#d0e0e3");
      
      // Default data untuk Settings
      if(sheetName === 'Settings') {
        sheet.appendRow(['SET1', 'H. Ahmad Fulan, S.Pd.I', '2024/2025']);
      }
    }
  }
}

// 2. MENGHADAPI REQUEST GET (Ambil Semua Data saat Web diload)
function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Fungsi pembantu mengubah baris data sheet menjadi Array JSON
  function getSheetData(sheetName) {
    const sheet = ss.getSheetByName(sheetName);
    if(!sheet) return [];
    const data = sheet.getDataRange().getValues();
    if(data.length <= 1) return [];
    
    const headers = data[0];
    const rows = data.slice(1);
    return rows.map(row => {
      let obj = {};
      headers.forEach((header, i) => obj[header] = row[i]);
      return obj;
    });
  }

  // Tarik data Settings khusus karena formatnya object, bukan array
  let settingsData = getSheetData('Settings')[0] || { Kepsek: '', TahunAktif: '' };

  let db = {
    settings: { kepsek: settingsData.Kepsek, tahun_aktif: settingsData.TahunAktif },
    kelas: getSheetData('Kelas'),
    cp_tp: getSheetData('CP_TP'),
    siswa: getSheetData('Siswa'),
    surat_masuk: getSheetData('SuratMasuk'),
    surat_keluar: getSheetData('SuratKeluar')
  };

  return ContentService.createTextOutput(JSON.stringify(db))
    .setMimeType(ContentService.MimeType.JSON);
}

// 3. MENGHADAPI REQUEST POST (Menyimpan Data dari Frontend)
function doPost(e) {
  let response = { success: false, message: 'Unknown action' };
  
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;
    const data = payload.data;
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (action === 'saveSiswa') {
      response = saveOrUpdateRow(ss.getSheetByName('Siswa'), data, 'id');
    } 
    else if (action === 'saveSettings') {
      const sheet = ss.getSheetByName('Settings');
      sheet.getRange(2, 2).setValue(data.kepsek);
      sheet.getRange(2, 3).setValue(data.tahun_aktif);
      response = { success: true, message: 'Settings saved' };
    }
    else if (action === 'saveCPTP') {
       // Pencarian berdasarkan kombinasi Mapel & Tingkat
       response = saveOrUpdateCPTP(ss.getSheetByName('CP_TP'), data);
    }
    else if (action === 'saveSuratKeluar') {
      response = saveOrUpdateRow(ss.getSheetByName('SuratKeluar'), data, 'id');
    }

    // Tambahkan aksi lain seperti saveKelas, saveNilai sesuai kebutuhan skala aplikasi

  } catch (error) {
    response = { success: false, message: error.toString() };
  }

  return ContentService.createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}

// Helper: Fungsi update baris jika ID ada, insert jika ID baru
function saveOrUpdateRow(sheet, dataObj, keyField) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  let rowIndex = -1;
  const keyIndex = headers.indexOf(keyField);
  
  // Cari apakah ID sudah ada
  for(let i=1; i<data.length; i++) {
    if(data[i][keyIndex] === dataObj[keyField]) {
      rowIndex = i + 1; // +1 karena sheet dimulai dari baris 1
      break;
    }
  }

  // Buat array row sesuai urutan header
  let newRow = headers.map(h => dataObj[h] !== undefined ? dataObj[h] : '');

  if(rowIndex > -1) {
    sheet.getRange(rowIndex, 1, 1, newRow.length).setValues([newRow]);
  } else {
    sheet.appendRow(newRow);
  }
  return { success: true, message: 'Data Tersimpan' };
}

// Helper: Update CP/TP berdasarkan Tingkat & Mapel
function saveOrUpdateCPTP(sheet, dataObj) {
   const data = sheet.getDataRange().getValues();
   const headers = data[0];
   let rowIndex = -1;
   
   const tIndex = headers.indexOf('tingkat');
   const mIndex = headers.indexOf('mapel');

   for(let i=1; i<data.length; i++) {
     if(data[i][tIndex] == dataObj.tingkat && data[i][mIndex] == dataObj.mapel) {
       rowIndex = i + 1; break;
     }
   }

   let newRow = headers.map(h => dataObj[h] !== undefined ? dataObj[h] : '');
   if(rowIndex > -1) {
     sheet.getRange(rowIndex, 1, 1, newRow.length).setValues([newRow]);
   } else {
     sheet.appendRow(newRow);
   }
   return { success: true, message: 'CP/TP Tersimpan' };
}
