/**
 * ============================================================
 * TUTUP_BUKU.GS — Proses Tutup Buku Tahunan
 * Koperasi App
 * Sesuai dokumen: REKAP_Koperasi_v2.md (BAGIAN 8, 12 #9, 16, 17)
 * ============================================================
 *
 * INI PROSES PALING "BERAT" DI SISTEM — dijalankan SEKALI SETAHUN
 * setelah RAT. Urutannya (sesuai BAGIAN 8 + 16):
 *
 * 1. VALIDASI PRA-TUTUP (bisa dicek dulu terpisah lewat menu):
 *    - Semua jasa sukarela tahun ini sudah POSTED (tidak ada DRAFT)
 *    - Jurnal balance (total debit = total kredit)
 *    - Setting RAT tahun BARU sudah dibuat di setting_rat
 * 2. Buat jurnal alokasi dana dari SHU (BAGIAN 12 #9):
 *    Debit 302 = Kredit 502 + 503 + 504
 * 3. LOCK semua transaksi tahun berjalan (status_lock → LOCKED)
 * 4. Hitung saldo akhir → tulis ke saldo_awal tahun baru
 * 5. Pindahkan status_aktif: tahun lama FALSE, tahun baru TRUE
 * 6. Generate semua laporan final (arsip)
 *
 * SETELAH TUTUP BUKU: input transaksi ke tahun lama otomatis
 * DITOLAK oleh validateTahunAktif (sudah terpasang di semua modul).
 * ============================================================
 */

// ============================================================
// VALIDASI PRA-TUTUP BUKU (BAGIAN 17)
// Bisa dijalankan terpisah untuk cek kesiapan tanpa eksekusi
// ============================================================

/**
 * Cek semua syarat tutup buku. Mengembalikan daftar masalah
 * (kosong = siap tutup buku).
 */
function cekKesiapanTutupBuku(tahunLama) {
  const masalah = [];
  const tahunBaru = Number(tahunLama) + 1;

  // 1. Setting tahun lama harus ada & aktif
  const settingLama = getSettingRAT(tahunLama);
  if (!settingLama) {
    masalah.push('Setting RAT tahun ' + tahunLama + ' tidak ditemukan.');
  } else if (settingLama.status_aktif !== true) {
    masalah.push('Tahun ' + tahunLama + ' tidak berstatus aktif — mungkin sudah ditutup?');
  }

  // 2. Setting tahun BARU harus sudah dibuat (hasil keputusan RAT)
  const settingBaru = getSettingRAT(tahunBaru);
  if (!settingBaru) {
    masalah.push('Setting RAT tahun ' + tahunBaru + ' BELUM dibuat. ' +
      'Tambahkan baris tahun ' + tahunBaru + ' di sheet setting_rat ' +
      '(isi sesuai keputusan RAT, status_aktif = FALSE dulu — ' +
      'nanti diaktifkan otomatis oleh proses tutup buku).');
  } else {
    try { validateAlokasiSHU(settingBaru); }
    catch (e) { masalah.push('Setting tahun ' + tahunBaru + ': ' + e.message); }
  }
  if (settingLama) {
    try { validateAlokasiSHU(settingLama); }
    catch (e) { masalah.push('Setting tahun ' + tahunLama + ': ' + e.message); }
  }

  // 3. Tidak boleh ada jasa sukarela DRAFT
  const draftJasa = getRowsByFilter(SHEET.JASA_SUKARELA, function(row) {
    return Number(row.tahun) === Number(tahunLama) && row.status_posting === 'DRAFT';
  });
  if (draftJasa.length > 0) {
    masalah.push('Masih ada ' + draftJasa.length + ' baris jasa sukarela berstatus DRAFT. ' +
      'Posting dulu (atau hapus jika salah) sebelum tutup buku.');
  }

  // 4. Jurnal harus balance
  const saldo = getSaldoPerAkun(tahunLama);
  let totalDebit = 0, totalKredit = 0;
  Object.keys(saldo).forEach(function(kode) {
    totalDebit += saldo[kode].debit;
    totalKredit += saldo[kode].kredit;
  });
  if (Math.abs(totalDebit - totalKredit) > 1) {
    masalah.push('Jurnal tahun ' + tahunLama + ' TIDAK BALANCE. Debit=' +
      formatRupiah(totalDebit) + ', Kredit=' + formatRupiah(totalKredit) +
      '. Telusuri jurnal_umum sebelum tutup buku.');
  }

  return masalah;
}

/** Menu: cek kesiapan tanpa mengeksekusi apapun. */
function menuCekKesiapanTutupBuku() {
  const tahun = getTahunAktif();
  if (!tahun) { tampilkanPesan('Tidak ada tahun aktif di setting_rat.'); return; }

  const masalah = cekKesiapanTutupBuku(tahun);
  tampilkanPesan(masalah.length === 0
    ? '✅ SIAP TUTUP BUKU tahun ' + tahun + '.\nSemua syarat terpenuhi.'
    : '⚠️ BELUM SIAP tutup buku tahun ' + tahun + ':\n\n• ' + masalah.join('\n• '));
}

// ============================================================
// PROSES TUTUP BUKU
// ============================================================

/**
 * Eksekusi tutup buku tahun aktif.
 * JANGAN panggil langsung — gunakan menuTutupBukuTahun yang
 * ada konfirmasi bertingkat.
 */
function prosesTutupBuku(tahunLama) {
  const tahunBaru = Number(tahunLama) + 1;

  // ---------- 0. VALIDASI FINAL ----------
  const masalah = cekKesiapanTutupBuku(tahunLama);
  if (masalah.length > 0) {
    throw new Error('Tutup buku dibatalkan:\n• ' + masalah.join('\n• '));
  }

  const laporan = [];

  // ---------- 1. JURNAL PENUTUP + ALOKASI SHU (7 pos) ----------
  // Pendapatan & biaya tahun lama dinolkan; SHU dipindah ke Cadangan
  // (modal sendiri) & dana-dana (hutang lancar) sesuai persen alokasi.
  const alokasi = buatJurnalPenutup(tahunLama);
  if (alokasi.shuKotor > 0) {
    laporan.push('Jurnal penutup dibuat. SHU ' + formatRupiah(alokasi.shuKotor) + ' dialokasikan: ' +
      alokasi.pos.map(function(p) { return p.label + ' ' + formatRupiah(p.nominal); }).join(', ') + '.');
  } else if (alokasi.shuKotor < 0) {
    laporan.push('Jurnal penutup dibuat. Rugi ' + formatRupiah(-alokasi.shuKotor) +
      ' mengurangi Cadangan.');
  } else {
    laporan.push('SHU = 0 → jurnal penutup tanpa alokasi.');
  }

  // ---------- 2. SALDO ANGGOTA ----------
  // Tidak disalin ke saldo_awal: saldo simpanan & piutang dihitung
  // kumulatif dari seluruh transaksi lintas tahun. (Versi lama menyalin
  // saldo sehingga simpanan terhitung dua kali setelah tutup buku.)

  // ---------- 3. LOCK SEMUA TRANSAKSI TAHUN LAMA ----------
  const sheetTransaksi = [SHEET.SIMPANAN, SHEET.PENGAMBILAN, SHEET.PINJAMAN,
                          SHEET.ANGSURAN, SHEET.KAS];
  let totalDikunci = 0;

  sheetTransaksi.forEach(function(namaSheet) {
    const sheet = getSheet(namaSheet);
    const header = getHeader(namaSheet);
    const idxTahun = header.indexOf('tahun');
    const idxLock = header.indexOf('status_lock');
    if (idxTahun === -1 || idxLock === -1) return;

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    // Baca & tulis kolom lock sekaligus (cepat, bukan per sel)
    const dataTahun = sheet.getRange(2, idxTahun + 1, lastRow - 1, 1).getValues();
    const dataLock = sheet.getRange(2, idxLock + 1, lastRow - 1, 1).getValues();

    let berubah = 0;
    for (let i = 0; i < dataTahun.length; i++) {
      if (Number(dataTahun[i][0]) === Number(tahunLama) &&
          dataLock[i][0] === STATUS_LOCK.OPEN) {
        dataLock[i][0] = STATUS_LOCK.LOCKED;
        berubah++;
      }
    }
    if (berubah > 0) {
      sheet.getRange(2, idxLock + 1, lastRow - 1, 1).setValues(dataLock);
      hapusCache(namaSheet);
      totalDikunci += berubah;
    }
  });
  laporan.push(totalDikunci + ' transaksi tahun ' + tahunLama + ' dikunci (LOCKED).');

  // ---------- 4. PINDAHKAN STATUS AKTIF ----------
  updateRowByField(SHEET.SETTING_RAT, 'tahun', tahunLama, { status_aktif: false });
  updateRowByField(SHEET.SETTING_RAT, 'tahun', tahunBaru, { status_aktif: true });
  laporan.push('Tahun aktif dipindah: ' + tahunLama + ' → ' + tahunBaru + '.');

  // ---------- 5. GENERATE LAPORAN FINAL (ARSIP) ----------
  const modeSebelumnya = MODE_SENYAP;
  MODE_SENYAP = true;
  generateLaporanLabaRugi(tahunLama);
  generateLaporanNeraca(tahunLama);
  generateLaporanSHU(tahunLama);
  MODE_SENYAP = modeSebelumnya;
  laporan.push('Laporan final tahun ' + tahunLama + ' digenerate (lap_labarugi, lap_neraca, lap_shu).');

  // ---------- 6. LOG ----------
  logAktivitas('EDIT', SHEET.SETTING_RAT, 'TUTUP-BUKU-' + tahunLama,
    { tahun_aktif: tahunLama }, { tahun_aktif: tahunBaru, detail: laporan.join(' | ') });

  return laporan;
}

// ============================================================
// MENU DENGAN KONFIRMASI BERTINGKAT
// ============================================================

function menuTutupBukuTahun() {
  const ui = SpreadsheetApp.getUi();
  const tahun = getTahunAktif();
  if (!tahun) { ui.alert('Tidak ada tahun aktif di setting_rat.'); return; }

  // Cek kesiapan dulu, tampilkan hasilnya
  const masalah = cekKesiapanTutupBuku(tahun);
  if (masalah.length > 0) {
    ui.alert('⚠️ BELUM SIAP tutup buku tahun ' + tahun + ':\n\n• ' +
      masalah.join('\n• ') + '\n\nSelesaikan dulu, lalu coba lagi.');
    return;
  }

  const konfirmasi1 = ui.alert(
    '⚠️ TUTUP BUKU TAHUN ' + tahun,
    'Proses ini akan:\n' +
    '1. Membuat jurnal alokasi dana dari SHU\n' +
    '2. MENGUNCI semua transaksi tahun ' + tahun + ' (tidak bisa diedit lagi)\n' +
    '3. Membuat saldo_awal tahun ' + (tahun + 1) + '\n' +
    '4. Memindahkan tahun aktif ke ' + (tahun + 1) + '\n' +
    '5. Meng-generate laporan final\n\n' +
    'Proses TIDAK BISA di-undo otomatis. Lanjutkan?',
    ui.ButtonSet.YES_NO
  );
  if (konfirmasi1 !== ui.Button.YES) return;

  const konfirmasi2 = ui.prompt(
    '⚠️ KONFIRMASI TERAKHIR',
    'Ketik angka tahun yang mau ditutup (' + tahun + ') untuk melanjutkan:',
    ui.ButtonSet.OK_CANCEL
  );
  if (konfirmasi2.getSelectedButton() !== ui.Button.OK) return;
  if (konfirmasi2.getResponseText().trim() !== String(tahun)) {
    ui.alert('Angka tahun tidak cocok. Tutup buku dibatalkan.');
    return;
  }

  try {
    const laporan = prosesTutupBuku(tahun);
    ui.alert('✅ TUTUP BUKU TAHUN ' + tahun + ' SELESAI\n\n• ' + laporan.join('\n• ') +
      '\n\nTahun ' + (tahun + 1) + ' sekarang aktif. Selamat tahun buku baru! 🎉');
  } catch (e) {
    ui.alert('❌ GAGAL: ' + e.message);
  }
}
// ============================================================
// TEST DARI EDITOR (tanpa parameter, tanpa dialog konfirmasi)
// ============================================================

/**
 * Cek kesiapan tutup buku dari editor — TIDAK mengeksekusi apapun,
 * hanya menampilkan daftar syarat yang belum terpenuhi di log.
 */
function test_cekKesiapanTutupBuku() {
  MODE_SENYAP = true;
  const tahun = getTahunAktif();
  if (!tahun) { Logger.log('❌ Tidak ada tahun aktif di setting_rat.'); return; }

  const masalah = cekKesiapanTutupBuku(tahun);
  Logger.log(masalah.length === 0
    ? '✅ SIAP tutup buku tahun ' + tahun
    : '⚠️ BELUM SIAP tahun ' + tahun + ':\n• ' + masalah.join('\n• '));
}

/**
 * Eksekusi tutup buku dari editor untuk TAHUN AKTIF saat ini.
 * ⚠️ TANPA konfirmasi — hanya untuk testing di spreadsheet dummy!
 * Di penggunaan asli, SELALU pakai menu 📕 TUTUP BUKU TAHUN.
 */
function test_prosesTutupBukuTahunAktif() {
  MODE_SENYAP = true;
  const tahun = getTahunAktif();
  if (!tahun) { Logger.log('❌ Tidak ada tahun aktif.'); return; }

  Logger.log('Menutup buku tahun ' + tahun + '...');
  try {
    const laporan = prosesTutupBuku(tahun);
    Logger.log('✅ SELESAI:\n• ' + laporan.join('\n• '));
    Logger.log('Tahun aktif sekarang: ' + getTahunAktif());
  } catch (e) {
    Logger.log('❌ ' + e.message);
  }
}