/**
 * ============================================================
 * TEST_LAPORAN.GS — Uji Coba Modul Laporan
 * ============================================================
 */

/** Tes: SHU kotor dari jurnal */
function test_getTotalSHUTahun() {
  MODE_SENYAP = true;
  const tahun = getTahunAktif();
  const shu = getTotalSHUTahun(tahun);
  Logger.log('SHU Kotor tahun ' + tahun + ': ' + formatRupiah(shu));
}

/** Tes: alokasi dana dari SHU */
function test_hitungAlokasiDana() {
  MODE_SENYAP = true;
  const tahun = getTahunAktif();
  const alokasi = hitungAlokasiDana(tahun);
  Logger.log(JSON.stringify(alokasi, null, 2));

  const cek = alokasi.danaCadangan + alokasi.danaPengurus +
              alokasi.danaSosial + alokasi.shuNetto;
  Logger.log('Cek: dana + netto = ' + formatRupiah(cek) +
    ' vs SHU kotor = ' + formatRupiah(alokasi.shuKotor) +
    (Math.abs(cek - alokasi.shuKotor) <= 2 ? ' ✅ (selisih pembulatan wajar)' : ' ❌'));
}

/** Tes: SHU satu anggota */
function test_hitungSHUAnggota() {
  MODE_SENYAP = true;
  const tahun = getTahunAktif();
  const anggotaAktif = getRowsByFilter(SHEET.ANGGOTA, function(r) {
    return r.status === 'aktif';
  });
  if (anggotaAktif.length === 0) { Logger.log('Tidak ada anggota aktif.'); return; }

  const target = anggotaAktif[0];
  const shu = hitungSHU(target.id_anggota, tahun);
  Logger.log('SHU ' + target.nama + ': ' + JSON.stringify(shu, null, 2));
}

/** Tes: total SHU semua anggota tidak boleh melebihi SHU netto */
function test_totalSHUTidakMelebihiNetto() {
  MODE_SENYAP = true;
  const tahun = getTahunAktif();
  const alokasi = hitungAlokasiDana(tahun);

  const anggotaAktif = getRowsByFilter(SHEET.ANGGOTA, function(r) {
    return r.status === 'aktif';
  });

  let total = 0;
  anggotaAktif.forEach(function(a) {
    total += hitungSHU(a.id_anggota, tahun).total;
  });

  Logger.log('SHU Netto: ' + formatRupiah(alokasi.shuNetto));
  Logger.log('Total SHU terbagi ke ' + anggotaAktif.length + ' anggota: ' + formatRupiah(total));
  // toleransi pembulatan: 1 rupiah per anggota
  const toleransi = anggotaAktif.length + 2;
  Logger.log(Math.abs(total - alokasi.shuNetto) <= toleransi
    ? '✅ Sesuai (selisih hanya pembulatan)'
    : '❌ Selisih terlalu besar: ' + formatRupiah(total - alokasi.shuNetto));
}

/** Tes: generate semua laporan sekaligus */
function test_generateSemuaLaporan() {
  MODE_SENYAP = true;
  generateLaporanLabaRugi();
  generateLaporanNeraca();
  generateLaporanSHU();
  generateRekapSimpanan();
  generateRekapPinjaman();
  Logger.log('✅ Semua laporan digenerate. Cek 5 sheet laporan untuk verifikasi visual.');
}

/** Jalankan semua tes laporan */
function test_semuaFungsiLaporan() {
  MODE_SENYAP = true;
  Logger.log('===== MULAI TES MODUL LAPORAN =====');
  Logger.log('\n--- 1. SHU kotor ---');
  test_getTotalSHUTahun();
  Logger.log('\n--- 2. Alokasi dana ---');
  test_hitungAlokasiDana();
  Logger.log('\n--- 3. SHU satu anggota ---');
  test_hitungSHUAnggota();
  Logger.log('\n--- 4. Total SHU vs netto ---');
  test_totalSHUTidakMelebihiNetto();
  Logger.log('\n--- 5. Generate semua laporan ---');
  test_generateSemuaLaporan();
  Logger.log('\n===== SELESAI =====');
}