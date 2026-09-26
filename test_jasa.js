/**
 * ============================================================
 * TEST_JASA.GS — Uji Coba Modul Jasa Sukarela
 * ============================================================
 * SYARAT: minimal 1 anggota aktif punya saldo simpanan sukarela
 * lebih dari 0 (dari saldo_awal migrasi ATAU dari setoran baru).
 * ============================================================
 */

/** Tes: hitung jasa sukarela bulan berjalan (DRAFT) */
function test_hitungJasaSukarelaBulanan() {
  const hasil = hitungJasaSukarelaBulanan();
  Logger.log('Dibuat: ' + hasil.dibuat.length);
  Logger.log('Dilewati (saldo 0): ' + hasil.dilewatiSaldoNol.length);
  Logger.log('Dilewati (sudah ada): ' + hasil.dilewatiSudahAda.length);
  Logger.log(JSON.stringify(hasil, null, 2));
}

/** Tes: hitung dobel di bulan sama — harus semua dilewati */
function test_hitungJasaDobel_harusDilewati() {
  const hasilPertama = hitungJasaSukarelaBulanan();
  const hasilKedua = hitungJasaSukarelaBulanan();
  Logger.log('Hitung pertama — dibuat: ' + hasilPertama.dibuat.length);
  Logger.log('Hitung kedua (harus 0 dibuat, semua dilewati) — dibuat: ' +
    hasilKedua.dibuat.length + ', dilewati sudah ada: ' + hasilKedua.dilewatiSudahAda.length);

  if (hasilKedua.dibuat.length === 0) {
    Logger.log('✅ Berhasil: tidak ada dobel hitung.');
  } else {
    Logger.log('❌ MASALAH: ada baris yang dobel dihitung!');
  }
}

/** Tes: posting jasa sukarela (kunci + jurnal) */
function test_postingJasaSukarela() {
  // Pastikan ada DRAFT dulu
  hitungJasaSukarelaBulanan();

  const hasil = postingJasaSukarela();
  Logger.log('Diposting: ' + hasil.diposting);
  Logger.log('Total jasa: ' + formatRupiah(hasil.totalJasa || 0));
}

/** Tes: cek jurnal jasa sukarela balance setelah posting */
function test_cekJurnalJasaSukarela() {
  hitungJasaSukarelaBulanan();
  postingJasaSukarela();

  const tahun = getTahunAktif();
  const bulan = new Date().getMonth() + 1;

  const jasaBulanIni = getRowsByFilter(SHEET.JASA_SUKARELA, function(row) {
    return Number(row.tahun) === Number(tahun) && Number(row.bulan) === Number(bulan) &&
           row.status_posting === 'POSTED';
  });

  if (jasaBulanIni.length === 0) {
    Logger.log('Tidak ada jasa POSTED untuk dicek (mungkin semua anggota saldo sukarela 0).');
    return;
  }

  jasaBulanIni.forEach(function(row) {
    const jurnal = getRowsByFilter(SHEET.JURNAL, function(j) { return j.referensi === row.id; });
    const totalDebit = jurnal.reduce(function(s, j) { return s + Number(j.debit); }, 0);
    const totalKredit = jurnal.reduce(function(s, j) { return s + Number(j.kredit); }, 0);
    Logger.log(row.id + ' — Debit: ' + totalDebit + ', Kredit: ' + totalKredit +
      (totalDebit === totalKredit ? ' ✅ BALANCE' : ' ❌ TIDAK BALANCE'));
  });
}

/** Tes: pastikan posting menambah saldo simpanan sukarela anggota */
function test_saldoSukarelaBertambahSetelahPosting() {
  const anggotaAktif = getRowsByFilter(SHEET.ANGGOTA, function(row) {
    return row.status === 'aktif';
  });
  if (anggotaAktif.length === 0) {
    Logger.log('Tidak ada anggota aktif untuk dicek.');
    return;
  }
  const target = anggotaAktif[0];

  const saldoSebelum = getSaldoSimpanan(target.id_anggota, 'sukarela');
  Logger.log('Saldo sukarela ' + target.nama + ' SEBELUM posting: ' + formatRupiah(saldoSebelum));

  hitungJasaSukarelaBulanan();
  postingJasaSukarela();

  const saldoSesudah = getSaldoSimpanan(target.id_anggota, 'sukarela');
  Logger.log('Saldo sukarela ' + target.nama + ' SESUDAH posting: ' + formatRupiah(saldoSesudah));
  Logger.log(saldoSesudah >= saldoSebelum ? '✅ Saldo bertambah/tetap (wajar jika saldo awal 0)' :
    '❌ Saldo berkurang — ada yang salah!');
}

/** Jalankan semua tes jasa sukarela sekaligus */
function test_semuaFungsiJasa() {
  Logger.log('===== MULAI TES MODUL JASA SUKARELA =====');
  Logger.log('\n--- 1. Hitung jasa sukarela (DRAFT) ---');
  test_hitungJasaSukarelaBulanan();
  Logger.log('\n--- 2. Hitung dobel (harus dilewati) ---');
  test_hitungJasaDobel_harusDilewati();
  Logger.log('\n--- 3. Posting (kunci + jurnal) ---');
  test_postingJasaSukarela();
  Logger.log('\n--- 4. Cek jurnal balance ---');
  test_cekJurnalJasaSukarela();
  Logger.log('\n--- 5. Cek saldo sukarela bertambah ---');
  test_saldoSukarelaBertambahSetelahPosting();
  Logger.log('\n===== SELESAI. Cek sheet rekap_jasa_sukarela dan jurnal_umum ' +
    'untuk verifikasi visual. =====');
}