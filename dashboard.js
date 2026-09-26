/**
 * ============================================================
 * DASHBOARD.GS — Ringkasan Saldo Per Anggota
 * Koperasi App
 * ============================================================
 *
 * TUJUAN:
 * Sheet-sheet seperti saldo_awal, transaksi_simpanan, dll
 * sengaja dibuat format PANJANG (satu anggota = banyak baris)
 * supaya gampang diproses script. Tapi format ini agak
 * merepotkan kalau mau dibaca manusia langsung.
 *
 * Script ini membuat RINGKASAN di sheet DASHBOARD — satu baris
 * per anggota, kolom ke samping untuk tiap jenis saldo — mirip
 * Excel lama, tapi ANGKANYA DIHITUNG OTOMATIS dari data asli
 * (bukan diketik manual, jadi tidak akan pernah beda sendiri).
 *
 * CARA PAKAI:
 * Menu 🔧 SETUP → 📊 Refresh Dashboard
 * (atau jalankan fungsi refreshDashboard dari editor)
 *
 * Jalankan ulang kapan saja setelah ada transaksi baru untuk
 * memperbarui angkanya.
 * ============================================================
 */

const HEADER_DASHBOARD = [
  'id_anggota', 'nama', 'jabatan', 'status',
  'simpanan_pokok', 'simpanan_wajib', 'simpanan_sukarela',
  'total_simpanan', 'sisa_pinjaman', 'saldo_kas_koperasi'
];

function refreshDashboard() {
  const sheet = getSheet(SHEET.DASHBOARD);
  const anggotaList = sheetToObjects(SHEET.ANGGOTA);

  // Bersihkan isi lama (kecuali baris 1 label lama, akan ditimpa header baru)
  sheet.clear();

  // Judul & waktu update
  sheet.getRange('A1').setValue('📊 RINGKASAN SALDO ANGGOTA — Koperasi ' + getProfilKoperasi().nama_koperasi.replace(/"/g, ''))
    .setFontWeight('bold').setFontSize(14);
  sheet.getRange('A2').setValue('Diperbarui: ' +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd MMMM yyyy, HH:mm'))
    .setFontStyle('italic').setFontColor('#666666');

  // Saldo kas koperasi (satu angka untuk semua, bukan per-anggota)
  const saldoKas = getSaldoKas();
  sheet.getRange('A3').setValue('Saldo Kas Koperasi saat ini:')
    .setFontWeight('bold');
  sheet.getRange('B3').setValue(saldoKas)
    .setNumberFormat('Rp #,##0').setFontWeight('bold').setFontColor('#38761d');

  // Header tabel mulai baris 5
  const baseRow = 5;
  const headerRange = sheet.getRange(baseRow, 1, 1, HEADER_DASHBOARD.length);
  headerRange.setValues([HEADER_DASHBOARD]);
  headerRange.setFontWeight('bold').setFontColor('#ffffff')
    .setBackground('#4a86e8').setHorizontalAlignment('center');
  sheet.setFrozenRows(baseRow);

  if (anggotaList.length === 0) {
    sheet.getRange(baseRow + 1, 1).setValue('(Belum ada data anggota)')
      .setFontStyle('italic').setFontColor('#999999');
    formatKolom(sheet, baseRow);
    return;
  }

  // Hitung baris per anggota
  const baris = anggotaList.map(function(anggota) {
    const pokok    = getSaldoSimpanan(anggota.id_anggota, 'pokok');
    const wajib    = getSaldoSimpanan(anggota.id_anggota, 'wajib');
    const sukarela = getSaldoSimpanan(anggota.id_anggota, 'sukarela');
    const totalSimpanan = pokok + wajib + sukarela;
    const sisaPinjaman = getSisaPinjamanAnggota(anggota.id_anggota);

    return [
      anggota.id_anggota,
      anggota.nama,
      anggota.jabatan,
      anggota.status,
      pokok,
      wajib,
      sukarela,
      totalSimpanan,
      sisaPinjaman,
      '' // saldo_kas_koperasi hanya diisi di baris ringkasan atas, kosongkan per baris
    ];
  });

  sheet.getRange(baseRow + 1, 1, baris.length, HEADER_DASHBOARD.length).setValues(amanBaris_(baris));

  // Baris TOTAL di paling bawah
  const totalRow = baseRow + baris.length + 1;
  sheet.getRange(totalRow, 2).setValue('TOTAL').setFontWeight('bold');
  ['E', 'F', 'G', 'H', 'I'].forEach(function(kolom, i) {
    const kolomIdx = 5 + i; // E=5 (simpanan_pokok) s/d I=9 (sisa_pinjaman)
    const range = sheet.getRange(baseRow + 1, kolomIdx, baris.length, 1);
    sheet.getRange(totalRow, kolomIdx).setFormula(
      '=SUM(' + range.getA1Notation() + ')'
    ).setFontWeight('bold');
  });

  formatKolom(sheet, baseRow);

  SpreadsheetApp.flush();
  tampilkanPesan('✅ Dashboard diperbarui. ' + anggotaList.length + ' anggota ditampilkan.');
}

/** Rapikan lebar kolom & format angka Rupiah */
function formatKolom(sheet, baseRow) {
  sheet.setColumnWidth(1, 100);  // id_anggota
  sheet.setColumnWidth(2, 160);  // nama
  sheet.setColumnWidth(3, 90);   // jabatan
  sheet.setColumnWidth(4, 80);   // status
  for (let c = 5; c <= 9; c++) sheet.setColumnWidth(c, 140);

  const lastRow = sheet.getLastRow();
  if (lastRow > baseRow) {
    sheet.getRange(baseRow + 1, 5, lastRow - baseRow, 5)
      .setNumberFormat('Rp #,##0');
  }
}

/**
 * Total sisa pinjaman AKTIF milik satu anggota
 * (bisa lebih dari satu pinjaman aktif secara teori,
 * meski aturan bisnis membatasi 1 pinjaman aktif per anggota).
 */
function getSisaPinjamanAnggota(idAnggota) {
  const pinjamanList = getRowsByFilter(SHEET.PINJAMAN, function(row) {
    return row.id_anggota === idAnggota && row.status === 'aktif';
  });
  return pinjamanList.reduce(function(total, p) {
    return total + getSisaPinjaman(p.id_pinjaman);
  }, 0);
}