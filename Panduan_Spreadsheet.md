# Panduan Integrasi Form ke Google Spreadsheet

Untuk menghubungkan form HTML ini dengan Google Spreadsheet, ikuti langkah-langkah berikut:

## 1. Buat Google Spreadsheet Baru
1. Buka [Google Sheets](https://sheets.google.com) dan buat dokumen kosong baru.
2. Pada **Sheet 1**, buat header di baris pertama (A1 sampai H1):
   - Kolom A: `Timestamp`
   - Kolom B: `Kelas`
   - Kolom C: `Nama Lengkap`
   - Kolom D: `Jenis Setoran`
   - Kolom E: `Surat`
   - Kolom F: `Ayat Mulai`
   - Kolom G: `Ayat Sampai`
   - Kolom H: `Nilai`
   - Kolom I: `Catatan`

## 2. Buat Google Apps Script
1. Di menu Google Spreadsheet, klik **Extensions (Ekstensi)** > **Apps Script**.
2. Hapus semua kode yang ada di editor, lalu *copy-paste* kode berikut:

```javascript
const sheetName = 'Database';
const scriptProp = PropertiesService.getScriptProperties();

function initialSetup () {
  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  scriptProp.setProperty('key', activeSpreadsheet.getId());
}

function doPost (e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    const doc = SpreadsheetApp.openById(scriptProp.getProperty('key'));
    const sheet = doc.getSheetByName(sheetName);

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const nextRow = sheet.getLastRow() + 1;

    const newRow = headers.map(function(header) {
      if (header === 'Timestamp') {
        return new Date();
      } else {
        // e.parameter is an object containing form data
        return e.parameter[header] || e.parameter[header.toLowerCase()] || e.parameter[header.replace(/\s+/g, '')];
      }
    });

    // Handle form input names manually if needed
    const dataRow = [
      new Date(),
      e.parameter.kelas || '',
      e.parameter.namaLengkap || '',
      e.parameter.jenisSetoran || '',
      e.parameter.surat || '',
      e.parameter.ayatMulai || '',
      e.parameter.ayatSampai || '',
      e.parameter.nilai || '',
      e.parameter.catatan || ''
    ];

    sheet.getRange(nextRow, 1, 1, dataRow.length).setValues([dataRow]);

    return ContentService
      .createTextOutput(JSON.stringify({ 'result': 'success', 'row': nextRow }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  catch (e) {
    return ContentService
      .createTextOutput(JSON.stringify({ 'result': 'error', 'error': e }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  finally {
    lock.releaseLock();
  }
}
```

## 3. Deploy Script
1. Simpan script tersebut (klik ikon Save / Ctrl+S).
2. Klik tombol **Run** pada fungsi `initialSetup`. (Google mungkin meminta izin otorisasi akun. Lanjutkan dan izinkan).
3. Klik tombol **Deploy** di sudut kanan atas, pilih **New deployment**.
4. Pada tipe deployment (simbol gerigi), pilih **Web app**.
5. Isi formulir:
   - Description: `API Form Tahfidz` (atau bebas)
   - Execute as: **Me** (akun Anda)
   - Who has access: **Anyone** (Siapa saja)
6. Klik **Deploy**.
7. Salin **Web app URL** yang diberikan.

## 4. Hubungkan URL ke HTML
1. Buka file `index.html` yang telah dibuat.
2. Cari baris berikut (sekitar baris 330):
   ```javascript
   const GOOGLE_SCRIPT_URL = 'URL_APPS_SCRIPT_ANDA';
   ```
3. Ganti `'URL_APPS_SCRIPT_ANDA'` dengan URL Web App yang baru saja disalin. Contoh:
   ```javascript
   const GOOGLE_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycb.../exec';
   ```
4. Simpan file HTML, dan form sudah siap digunakan!
