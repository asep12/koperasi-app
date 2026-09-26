/**
 * ============================================================
 * TEST_SIMPANAN.GS — Uji Coba Modul Simpanan
 * ============================================================
 * Jalankan fungsi-fungsi ini SETELAH data anggota & saldo_awal
 * sudah terisi (butuh minimal 1 anggota aktif untuk tes).
 *
 * GANTI 'AGT-0001' di bawah dengan ID anggota yang benar-benar
 * ada di sheet anggota Anda sebelum menjalankan.
 * ============================================================
 */

/** Tes: input setoran simpanan wajib */
function test_inputSetoranWajib() {
  try {
    const hasil = inputSetoranSimpanan({
      id_anggota: 'AGT-0002',       // ⚠️ ganti sesuai ID anggota Anda
      jenis_simpanan: 'wajib',
      jumlah_setoran: 50000,
      keterangan: 'Tes setoran wajib bulan ini'
    });
    Logger.log('✅ ' + hasil.pesan);
    Logger.log('ID transaksi: ' + hasil.id_transaksi);
  } catch (e) {
    Logger.log('❌ ' + e.message);
  }
}

/** Tes: input setoran sukarela */
function test_inputSetoranSukarela() {
  try {
    const hasil = inputSetoranSimpanan({
      id_anggota: 'AGT-0002',       // ⚠️ ganti sesuai ID anggota Anda
      jenis_simpanan: 'sukarela',
      jumlah_setoran: 100000,
      keterangan: 'Tes setoran sukarela'
    });
    Logger.log('✅ ' + hasil.pesan);
  } catch (e) {
    Logger.log('❌ ' + e.message);
  }
}

/** Tes: setoran sukarela di bawah minimal — HARUS gagal */
function test_setoranSukarelaTerlaluKecil_gagal() {
  try {
    inputSetoranSimpanan({
      id_anggota: 'AGT-0002',
      jenis_simpanan: 'sukarela',
      jumlah_setoran: 5000  // di bawah minimal 10.000
    });
    Logger.log('⚠️ Seharusnya gagal, tapi malah berhasil!');
  } catch (e) {
    Logger.log('✅ Berhasil ditolak sesuai harapan: ' + e.message);
  }
}

/** Tes: setor ke tahun yang tidak aktif — HARUS gagal */
function test_setoranTahunTidakAktif_gagal() {
  try {
    inputSetoranSimpanan({
      id_anggota: 'AGT-0002',
      jenis_simpanan: 'sukarela',
      jumlah_setoran: 50000,
      tanggal: new Date(2020, 0, 1) // tahun 2020, pasti tidak ada di setting_rat
    });
    Logger.log('⚠️ Seharusnya gagal, tapi malah berhasil!');
  } catch (e) {
    Logger.log('✅ Berhasil ditolak sesuai harapan: ' + e.message);
  }
}

/** Tes: cek jurnal & kas otomatis terbentuk setelah setoran */
function test_cekJurnalOtomatis() {
  const hasil = inputSetoranSimpanan({
    id_anggota: 'AGT-0002',
    jenis_simpanan: 'wajib',
    jumlah_setoran: 50000,
    keterangan: 'Tes cek jurnal'
  });

  // Cek transaksi_kas
  const kas = getRowsByFilter(SHEET.KAS, function(row) {
    return row.referensi === hasil.id_transaksi;
  });
  Logger.log('Baris kas yang terbentuk: ' + JSON.stringify(kas));

  // Cek jurnal_umum
  const jurnal = getRowsByFilter(SHEET.JURNAL, function(row) {
    return row.referensi === hasil.id_transaksi;
  });
  Logger.log('Baris jurnal yang terbentuk: ' + JSON.stringify(jurnal));

  const totalDebit = jurnal.reduce(function(s, j) { return s + Number(j.debit); }, 0);
  const totalKredit = jurnal.reduce(function(s, j) { return s + Number(j.kredit); }, 0);
  Logger.log('Total Debit: ' + totalDebit + ' | Total Kredit: ' + totalKredit +
    (totalDebit === totalKredit ? ' ✅ BALANCE' : ' ❌ TIDAK BALANCE'));
}

/** Tes: pengajuan pengambilan sukarela (status PENDING) */
function test_inputPengajuanPengambilan() {
  try {
    const hasil = inputPengajuanPengambilan({
      id_anggota: 'AGT-0002',
      jumlah: 50000,
      keterangan: 'Tes pengajuan pengambilan'
    });
    Logger.log('✅ ' + hasil.pesan);
  } catch (e) {
    Logger.log('❌ ' + e.message);
  }
}

/** Jalankan semua tes simpanan sekaligus */
function test_semuaFungsiSimpanan() {
  Logger.log('===== MULAI TES MODUL SIMPANAN =====');
  Logger.log('\n--- 1. Setoran wajib normal ---');
  test_inputSetoranWajib();
  Logger.log('\n--- 2. Setoran sukarela normal ---');
  test_inputSetoranSukarela();
  Logger.log('\n--- 3. Setoran sukarela kekecilan (harus gagal) ---');
  test_setoranSukarelaTerlaluKecil_gagal();
  Logger.log('\n--- 4. Setor ke tahun tidak aktif (harus gagal) ---');
  test_setoranTahunTidakAktif_gagal();
  Logger.log('\n--- 5. Cek jurnal & kas otomatis ---');
  test_cekJurnalOtomatis();
  Logger.log('\n--- 6. Pengajuan pengambilan sukarela ---');
  test_inputPengajuanPengambilan();
  Logger.log('\n===== SELESAI. Cek juga langsung ke sheet transaksi_simpanan, ' +
    'transaksi_kas, dan jurnal_umum untuk verifikasi visual. =====');
}