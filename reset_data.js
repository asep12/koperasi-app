/**
 * ============================================================
 * RESET_DATA.GS — Bersihkan Data Uji Coba, Pertahankan Master
 * ============================================================
 *
 * BEDANYA DENGAN "Hapus Semua & Setup Ulang":
 * - Setup Ulang  = menghapus SEMUA sheet termasuk anggota (total)
 * - Reset Data   = hanya MENGOSONGKAN ISI sheet transaksi & laporan;
 *                  sheet-nya tetap ada (header, dropdown, warna utuh)
 *
 * YANG DIPERTAHANKAN (tidak disentuh sama sekali):
 * - anggota      (semua data anggota aman)
 * - coa_akun     (daftar akun)
 * - setting_rat  (baris tahun dipertahankan — lihat catatan bawah)
 *
 * YANG DIKOSONGKAN (isi dihapus, header tetap):
 * - saldo_awal, transaksi_simpanan, transaksi_pengambilan,
 *   transaksi_pinjaman, transaksi_angsuran, transaksi_kas,
 *   jurnal_umum, rekap_jasa_sukarela, realisasi_shu, log_aktivitas
 * - semua sheet laporan (lap_*, rekap_*, struk_potongan) + DASHBOARD
 *
 * CATATAN setting_rat: baris tahun TIDAK dihapus otomatis (keputusan
 * RAT itu data penting). Tapi setelah tes tutup buku, status_aktif
 * Anda mungkin sudah pindah ke tahun berikutnya — fungsi ini akan
 * MENAWARKAN mengembalikan status_aktif ke tahun yang Anda pilih.
 *
 * CARA PAKAI: Menu 🔧 SETUP → 🧹 Reset Data (Anggota Aman)
 * ============================================================
 */

// Sheet yang isinya dikosongkan (baris 2 ke bawah), header dipertahankan
const SHEET_DIKOSONGKAN = [
  'saldo_awal',
  'transaksi_simpanan',
  'transaksi_pengambilan',
  'transaksi_pinjaman',
  'transaksi_angsuran',
  'transaksi_kas',
  'jurnal_umum',
  'rekap_jasa_sukarela',
  'realisasi_shu',
  'log_aktivitas'
];

// Sheet laporan yang dibersihkan total (isinya memang generate-ulang)
const SHEET_LAPORAN_DIBERSIHKAN = [
  'DASHBOARD', 'lap_shu', 'lap_neraca', 'lap_labarugi',
  'rekap_pinjaman', 'rekap_angsuran', 'rekap_simpanan', 'struk_potongan'
];

/**
 * Fungsi inti reset. Bisa dipanggil dari menu atau test.
 * @param {number|null} tahunAktifKembali  jika diisi, status_aktif di
 *        setting_rat dikembalikan ke tahun ini (yang lain jadi FALSE)
 */
function resetDataTransaksi(tahunAktifKembali) {
  pastikanAksesData();
  const hasil = [];

  // 1. Kosongkan sheet transaksi (baris 2 ke bawah, header aman)
  SHEET_DIKOSONGKAN.forEach(function(nama) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nama);
    if (!sheet) { hasil.push('⏭️ ' + nama + ' tidak ditemukan, dilewati'); return; }

    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
      hasil.push('🧹 ' + nama + ': ' + (lastRow - 1) + ' baris dihapus');
    } else {
      hasil.push('✔️ ' + nama + ': sudah kosong');
    }
    hapusCache(nama);
  });

  // 2. Bersihkan sheet laporan (generate-ulang kapan saja)
  SHEET_LAPORAN_DIBERSIHKAN.forEach(function(nama) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nama);
    if (sheet) { sheet.clear(); hapusCache(nama); }
  });
  hasil.push('🧹 Semua sheet laporan & DASHBOARD dibersihkan');

  // 3. (Opsional) Kembalikan status_aktif setting_rat
  if (tahunAktifKembali) {
    const semuaSetting = sheetToObjects(SHEET.SETTING_RAT);
    semuaSetting.forEach(function(row) {
      const harusAktif = Number(row.tahun) === Number(tahunAktifKembali);
      if (row.status_aktif !== harusAktif) {
        updateRowByRowNumber(SHEET.SETTING_RAT, row.__row, { status_aktif: harusAktif });
      }
    });
    hasil.push('⚙️ status_aktif dikembalikan ke tahun ' + tahunAktifKembali);
  }

  hasil.push('');
  hasil.push('✅ DIPERTAHANKAN: anggota (' +
    sheetToObjects(SHEET.ANGGOTA).length + ' baris), coa_akun, setting_rat');

  return hasil;
}

// ============================================================
// MENU (dengan konfirmasi)
// ============================================================

function menuResetDataTransaksi() {
  const ui = SpreadsheetApp.getUi();

  const konfirmasi = ui.alert(
    '🧹 RESET DATA (Anggota Aman)',
    'Ini akan MENGHAPUS ISI semua sheet transaksi, jurnal, log, dan laporan — ' +
    'tapi data ANGGOTA, coa_akun, dan setting_rat TETAP AMAN.\n\n' +
    'Cocok untuk membersihkan data uji coba sebelum mulai input data asli.\n\n' +
    'Data transaksi yang terhapus TIDAK BISA dikembalikan. Lanjutkan?',
    ui.ButtonSet.YES_NO
  );
  if (konfirmasi !== ui.Button.YES) return;

  // Tawarkan kembalikan tahun aktif (berguna setelah tes tutup buku)
  const respTahun = ui.prompt(
    'Tahun aktif',
    'Setelah reset, tahun berapa yang jadi tahun aktif di setting_rat?\n' +
    '(Ketik angka tahun, contoh: 2026 — atau kosongkan untuk tidak mengubah)',
    ui.ButtonSet.OK_CANCEL
  );
  if (respTahun.getSelectedButton() !== ui.Button.OK) return;

  const teksTahun = respTahun.getResponseText().trim();
  let tahunAktif = null;
  if (teksTahun !== '') {
    tahunAktif = parseInt(teksTahun, 10);
    if (isNaN(tahunAktif) || !getSettingRAT(tahunAktif)) {
      ui.alert('Tahun "' + teksTahun + '" tidak ada di setting_rat. Reset dibatalkan — ' +
        'tidak ada data yang diubah.');
      return;
    }
  }

  const hasil = resetDataTransaksi(tahunAktif);
  ui.alert('✅ RESET SELESAI\n\n' + hasil.join('\n'));
}

// ============================================================
// VERSI EDITOR (untuk testing, tanpa dialog)
// ============================================================

/** Reset dari editor — tahun aktif dikembalikan ke 2026. Ubah jika perlu. */
function test_resetDataTransaksi() {
  MODE_SENYAP = true;
  const hasil = resetDataTransaksi(2026);  // ⚠️ ganti tahun jika perlu
  Logger.log(hasil.join('\n'));
}