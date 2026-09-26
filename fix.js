/**
 * ============================================================
 * FIX: Hapus opsi "kas" dan "shu" dari dropdown jenis saldo_awal
 * ============================================================
 * Alasan:
 * - "kas" bukan milik anggota tertentu, jadi tidak cocok di
 *   sheet yang barisnya per id_anggota.
 * - "shu" dihapus karena SHU koperasi selalu dibayar LUNAS
 *   setiap tahun, tidak pernah ada yang ditahan.
 *
 * Jenis yang tersisa & valid: simpanan_pokok, simpanan_wajib,
 * simpanan_sukarela, piutang.
 *
 * CARA PAKAI: pilih fungsi ini di dropdown editor, klik Run.
 * Aman dijalankan berkali-kali.
 * ============================================================
 */
function fix_hapusJenisKasShu() {
  const sheet = getSheet(SHEET.SALDO_AWAL);
  const header = getHeader(SHEET.SALDO_AWAL);
  const idxJenis = header.indexOf('jenis');

  if (idxJenis === -1) {
    Logger.log('❌ Kolom "jenis" tidak ditemukan di saldo_awal.');
    return;
  }

  // 1. Cek dulu apakah ada baris yang SUDAH terisi jenis kas/shu
  const data = sheetToObjects(SHEET.SALDO_AWAL);
  const bermasalah = data.filter(function(row) {
    return row.jenis === 'kas' || row.jenis === 'shu';
  });

  if (bermasalah.length > 0) {
    Logger.log('⚠️ PERHATIAN: ditemukan ' + bermasalah.length +
      ' baris dengan jenis kas/shu yang SUDAH TERISI:');
    bermasalah.forEach(function(row) {
      Logger.log('  Baris ' + row.__row + ': ' + row.id + ' — ' + row.jenis +
        ' — ' + formatRupiah(row.nominal));
    });
    Logger.log('Baris ini TIDAK dihapus otomatis — silakan cek & hapus manual ' +
      'jika memang keliru, sebelum melanjutkan.');
  } else {
    Logger.log('✅ Tidak ada baris lama yang pakai jenis kas/shu. Aman.');
  }

  // 2. Pasang ulang dropdown dengan daftar baru (tanpa kas, tanpa shu)
  const jenisBaru = ['simpanan_pokok', 'simpanan_wajib', 'simpanan_sukarela', 'piutang'];
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(jenisBaru, true)
    .setAllowInvalid(false)
    .setHelpText('Pilih: ' + jenisBaru.join(' / '))
    .build();

  sheet.getRange(2, idxJenis + 1, 1000, 1).setDataValidation(rule);

  // 3. Update catatan header supaya sesuai
  sheet.getRange(1, idxJenis + 1).setNote(
    'Jenis saldo yang dimigrasi dari Excel lama:\n' +
    '• simpanan_pokok    = saldo simpanan pokok\n' +
    '• simpanan_wajib    = akumulasi simpanan wajib\n' +
    '• simpanan_sukarela = saldo sukarela terakhir\n' +
    '• piutang = SISA pinjaman yang belum lunas\n' +
    '  (WAJIB dibuatkan juga baris pinjaman MIGRASI\n' +
    '   di transaksi_pinjaman — lihat catatan di sana)\n\n' +
    'CATATAN: jenis "kas" dan "shu" TIDAK dipakai —\n' +
    'kas bukan milik anggota tertentu, dan SHU koperasi\n' +
    'selalu dibayar lunas (tidak ada yang ditahan).'
  );

  Logger.log('✅ Dropdown & catatan jenis saldo_awal sudah diperbarui. ' +
    'Sekarang hanya ada 4 pilihan: ' + jenisBaru.join(', '));
}