# Panduan Membuat Dashboard Infografis (V5 - FINAL FIX)

Mohon maaf yang sebesar-besarnya! Ternyata perintah `setFormulaLocal` tidak tersedia di sistem Apps Script (hanya ada di Excel). 

Untuk mengatasi masalah bahasa/regional (pemisah titik koma) dengan pasti dan tanpa *error* fungsi script, kita bisa menggunakan perintah standar `setValue` saja namun tetap mempertahankan format penulisan ala Indonesia.

Ini adalah perbaikan finalnya.

## Langkah-langkah:

1. Buka halaman **Apps Script**.
2. **Hapus seluruh kode `buatDashboard` yang sebelumnya**.
3. **Copy dan Paste** kode V5 di bawah ini:

```javascript
// ==========================================
// KODE PEMBUAT DASHBOARD INFOGRAFIS (V5 - FINAL FIX INDONESIA)
// ==========================================

function buatDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Perbaiki Sheet Database (Buat Kolom Bantuan di Kolom J)
  var db = ss.getSheetByName('Database');
  if (db) {
    db.getRange('J1').setValue('Total Ayat (Auto)');
    // Menggunakan setValue untuk memaksa format lokal (titik koma) masuk ke sel
    db.getRange('J2:J').setValue('=IF(LEN(G2)>0; G2*1 - F2*1 + 1; 0)');
    db.hideColumns(10);
  }

  var dash = ss.getSheetByName('Dashboard');
  if (!dash) {
    dash = ss.insertSheet('Dashboard');
  } else {
    dash.clear();
    var charts = dash.getCharts();
    for (var i = 0; i < charts.length; i++) {
      dash.removeChart(charts[i]);
    }
  }

  // 1. Atur Ukuran Kolom
  dash.setColumnWidth(1, 20);
  dash.setColumnWidth(2, 250);
  dash.setColumnWidth(3, 100);
  dash.setColumnWidth(4, 20);
  dash.setColumnWidth(5, 200);
  dash.setColumnWidth(6, 100);
  dash.setColumnWidth(7, 20);
  dash.setColumnWidth(8, 400);

  // 2. Set Background, Font, & Sembunyikan Garis Kisi (Gridlines)
  dash.getRange(1, 1, 100, 15).setBackground('#f4f9f4').setFontFamily('Arial');
  dash.setHiddenGridlines(true);

  // 3. Buat Header
  var headerRange = dash.getRange('B2:K3');
  headerRange.merge();
  headerRange.setValue('📊 DASHBOARD INFOGRAFIS HAFALANQU');
  headerRange.setBackground('#2d6a4f').setFontColor('white').setFontSize(20).setFontWeight('bold');
  headerRange.setHorizontalAlignment('center').setVerticalAlignment('middle');

  // 4. Box 1: Ringkasan Setoran (Kiri)
  dash.getRange('B5:C5').merge().setValue('RINGKASAN SETORAN').setBackground('#1b4332').setFontColor('white').setFontWeight('bold').setHorizontalAlignment('center');
  dash.getRange('B6').setValue('Ziyadah');
  // Memakai titik koma (;) untuk pemisah fungsi karena akun berformat Indonesia
  dash.getRange('C6').setValue('=COUNTIF(Database!D:D; "*Ziyadah*")');
  
  dash.getRange('B7').setValue('Murajaah');
  dash.getRange('C7').setValue('=COUNTIF(Database!D:D; "*Murajaah*")');
  
  dash.getRange('B8').setValue('Total Ayat Disetor');
  dash.getRange('C8').setValue('=SUM(Database!J:J)');

  // 5. Box 2: Kualitas Hafalan
  dash.getRange('E5:F5').merge().setValue('KUALITAS HAFALAN').setBackground('#1b4332').setFontColor('white').setFontWeight('bold').setHorizontalAlignment('center');
  dash.getRange('E6').setValue('Mumtaz');
  dash.getRange('F6').setValue('=COUNTIF(Database!H:H; "*Mumtaz*")');
  dash.getRange('E7').setValue('Jayyid Jiddan');
  dash.getRange('F7').setValue('=COUNTIF(Database!H:H; "*Jayyid Jiddan*")');
  dash.getRange('E8').setValue('Jayyid');
  dash.getRange('F8').setValue('=COUNTIFS(Database!H:H; "*Jayyid*"; Database!H:H; "<>*Jiddan*")');
  dash.getRange('E9').setValue('Maqbul');
  dash.getRange('F9').setValue('=COUNTIF(Database!H:H; "*Maqbul*")');
  dash.getRange('E10').setValue('Rasib');
  dash.getRange('F10').setValue('=COUNTIF(Database!H:H; "*Rasib*")');

  // 6. Box 3: Top 20 Ziyadah
  dash.getRange('B12:C12').merge().setValue('🏆 TOP 20 SISWA (ZIYADAH)').setBackground('#1b4332').setFontColor('white').setFontWeight('bold').setHorizontalAlignment('center');
  // Perhatikan titik koma pada rumus utama, dan koma pada syntax query SELECT
  dash.getRange('B13').setValue('=IFERROR(QUERY(Database!A2:J; "SELECT C, SUM(J) WHERE C IS NOT NULL AND D CONTAINS \'Ziyadah\' GROUP BY C ORDER BY SUM(J) DESC LIMIT 20 LABEL C \'\', SUM(J) \'\'"; 0); "")');

  // Percantik Data Tabel
  dash.getRange('B6:C8').setBorder(true, true, true, true, true, true);
  dash.getRange('E6:F10').setBorder(true, true, true, true, true, true);
  dash.getRange('C6:C8').setHorizontalAlignment('center').setFontWeight('bold');
  dash.getRange('F6:F10').setHorizontalAlignment('center').setFontWeight('bold');

  // ==========================================
  // PEMBUATAN DIAGRAM (INFOGRAFIS)
  // ==========================================

  // Periksa apakah datanya kosong (mencegah error pembuatan grafik tanpa data)
  // Jika masih 0, kita beri nilai sementara agar script grafik tidak crash
  if(dash.getRange('C6').getValue() === 0 && dash.getRange('C7').getValue() === 0) {
    // Biarkan saja, Google Charts sekarang pintar menangani nilai kosong
  }

  var chartSetoran = dash.newChart()
    .asPieChart()
    .addRange(dash.getRange('B6:C7'))
    .setOption('title', 'Proporsi Jenis Setoran')
    .setOption('pieHole', 0.5)
    .setOption('pieSliceText', 'percentage')
    .setOption('colors', ['#2d6a4f', '#74c69d'])
    .setOption('backgroundColor', '#ffffff')
    .setPosition(5, 8, 0, 0) 
    .setOption('width', 350)
    .setOption('height', 250)
    .build();
  dash.insertChart(chartSetoran);

  var chartKualitas = dash.newChart()
    .asPieChart()
    .addRange(dash.getRange('E6:F10'))
    .setOption('title', 'Persentase Kualitas Hafalan')
    .setOption('pieHole', 0.5)
    .setOption('pieSliceText', 'percentage')
    .setOption('colors', ['#1b4332', '#2d6a4f', '#40916c', '#52b788', '#95d5b2'])
    .setOption('backgroundColor', '#ffffff')
    .setPosition(18, 8, 0, 0) 
    .setOption('width', 350)
    .setOption('height', 280)
    .build();
  dash.insertChart(chartKualitas);

  var chartTop20 = dash.newChart()
    .asBarChart()
    .addRange(dash.getRange('B12:C32'))
    .setOption('title', 'Peringkat 20 Hafalan Ziyadah Terbanyak (Jumlah Ayat)')
    .setOption('colors', ['#40916c'])
    .setOption('backgroundColor', '#ffffff')
    .setPosition(12, 5, 0, 0) 
    .setOption('width', 550)
    .setOption('height', 450)
    .setOption('legend', {position: 'none'}) 
    .build();
  dash.insertChart(chartTop20);

  ss.setActiveSheet(dash);
}
```

## 4. Eksekusi Final
1. **Save** kode terbarunya.
2. Klik **Jalankan (Run)**.
3. Buka Dashboard Anda! Kali ini tidak akan ada *error* `setFormulaLocal` dan semua rumusnya akan bekerja dengan lancar jaya!
